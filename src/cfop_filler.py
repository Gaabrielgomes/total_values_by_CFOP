import pandas as pd
import os
import re
from pathlib import Path
from tkinter import Tk, StringVar, Frame, Label, Entry, Button, filedialog, messagebox
from connection import connect_database
from query_loader import get_query


def validate_filename(name):
    forbidden_chars = r'[<>:"/\\|?*\[\]{}()+,;=\s]'
    pattern = re.compile(forbidden_chars)
    return name.strip() and not pattern.search(name)


def check_write_permission(path):
    try:
        test_file = path.parent / f"teste_permissao_{os.getpid()}.tmp"
        test_file.touch(exist_ok=True)
        test_file.unlink()
        return True
    except Exception:
        return False


def get_save_path():
    file_types = [('Excel', '*.xlsx'), ('Todos os arquivos', '*.*')]

    while True:
        file_path = filedialog.asksaveasfilename(
            title="Salvar relatório CFOP",
            defaultextension=".xlsx",
            filetypes=file_types,
            initialfile="relatorio_cfop.xlsx"
        )

        if not file_path:
            return None

        file_path = Path(file_path)

        if not validate_filename(file_path.name):
            messagebox.showerror("Erro", "Nome de arquivo inválido.")
            continue

        if not check_write_permission(file_path):
            messagebox.showerror("Erro", "Sem permissão de escrita no diretório selecionado.")
            continue

        return file_path


def process_db_results(results) -> dict:
    data = {}
    for row in results:
        codi_emp, codi_nat = str(row[0]), str(row[1])
        value = float(row[2]) if row[2] is not None else 0.0
        pis = float(row[3]) if row[3] is not None else 0.0
        cofins = float(row[4]) if row[4] is not None else 0.0
        icms = float(row[5]) if row[5] is not None else 0.0

        if codi_emp not in data:
            data[codi_emp] = {
                "COMPRAS": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "D_COMPRAS": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "CONSUMO": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "VENDAS": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "D_VENDAS": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "TRANSF_EN": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "TRANSF_SAI": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "OUTRAS_EN": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "OUTRAS_SAI": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0},
                "SERVICOS": {"VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0}
            }

        def update(cat):
            data[codi_emp][cat]["VALOR"] += value
            data[codi_emp][cat]["PIS"] += pis
            data[codi_emp][cat]["COFINS"] += cofins
            data[codi_emp][cat]["ICMS"] += icms

        if codi_nat in ['1102', '2102', '1253', '1403', '2403'] and str(row[6]) == 'ENTRADA':
            update("COMPRAS")
        elif codi_nat in ['5202', '6202', '6411'] and str(row[6]) == 'SAIDA':
            update("D_COMPRAS")
        elif codi_nat in ['1556', '2556', '1407', '2407'] and str(row[6]) == 'ENTRADA':
            update("CONSUMO")
        elif codi_nat in ['5101', '6101', '5102', '6102', '6108', '5402', '6402', '5403', '6403', '5404', '6404', '5405', '6405'] and str(row[6]) == 'SAIDA':
            update("VENDAS")
        elif codi_nat in ['1202', '2202', '1411', '2411'] and str(row[6]) == 'ENTRADA':
            update("D_VENDAS")
        elif codi_nat in ['1152', '2152', '1409', '2409'] and str(row[6]) == 'ENTRADA':
            update("TRANSF_EN")
        elif codi_nat in ['5152', '6152', '5409', '6409'] and str(row[6]) == 'SAIDA':
            update("TRANSF_SAI")
        elif codi_nat in ['1908', '1912', '1913', '2303', '1303', '1949', '2353', '2949', '2910'] and str(row[6]) == 'ENTRADA':
            update("OUTRAS_EN")
        elif codi_nat in ['5910', '5912', '5927', '5929', '6949'] and str(row[6]) == 'SAIDA':
            update("OUTRAS_SAI")
        elif codi_nat in ['1933', '2933'] and str(row[6]) == 'ENTRADA':
            update("SERVICOS")
    return data


def execute_query(companies_str, start_date, end_date, root):
    if not all([companies_str, start_date, end_date]):
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    try:
        company_list = [c.strip() for c in companies_str.split(',')]
        placeholders = ','.join('?' for _ in company_list)

        query_template = get_query("totalizer")
        query = query_template.replace("IN ({})", f"IN ({placeholders})")

        conn = connect_database()
        cursor = conn.cursor()

        params = (
            *company_list, start_date, end_date,
            *company_list, start_date, end_date,
            *company_list, start_date, end_date,
            *company_list, start_date, end_date
        )

        cursor.execute(query, params)
        results = cursor.fetchall()
        conn.close()

        info_dict = process_db_results(results)

        if not info_dict:
            messagebox.showinfo("Informação", "Nenhum dado encontrado.")
            return

        category_order = [
            "COMPRAS", "D_COMPRAS", "CONSUMO", "VENDAS", "D_VENDAS",
            "TRANSF_EN", "TRANSF_SAI", "OUTRAS_EN", "OUTRAS_SAI", "SERVICOS"
        ]

        records = []
        first_company = True

        for company, categories in sorted(info_dict.items(), key=lambda x: int(x[0])):

            if not first_company:
                records.append({col: "" for col in ["EMPRESA", "TIPO"] + category_order})
            first_company = False

            value_row = {"EMPRESA": company, "TIPO": "VALOR"}
            pis_row = {"EMPRESA": "", "TIPO": "PIS"}
            cofins_row = {"EMPRESA": "", "TIPO": "COFINS"}
            icms_row = {"EMPRESA": "", "TIPO": "ICMS"}

            for category in category_order:
                values = categories.get(category, {
                    "VALOR": 0, "PIS": 0, "COFINS": 0, "ICMS": 0
                })

                value_row[category] = values["VALOR"]
                pis_row[category] = values["PIS"]
                cofins_row[category] = values["COFINS"]
                icms_row[category] = values["ICMS"]

            records.extend([value_row, pis_row, cofins_row, icms_row])

        df = pd.DataFrame(records)
        df = df[["EMPRESA", "TIPO"] + category_order]

        save_path = get_save_path()
        if not save_path:
            return

        df.to_excel(save_path, index=False)
        messagebox.showinfo("Sucesso", f"Salvo em:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))


def window():
    root = Tk()
    root.title("Totalizar por CFOP")
    root.geometry(
        f"350x200+{(root.winfo_screenwidth() - 350) // 2}+{(root.winfo_screenheight() - 200) // 2}")
    root.resizable(True, True)

    main_frame = Frame(root, padx=20, pady=20)
    main_frame.pack(fill="both", expand=True)

    companies_var = StringVar()
    start_date_var = StringVar()
    end_date_var = StringVar()

    Label(main_frame, text="Empresas:").pack(anchor="w")
    Entry(main_frame, textvariable=companies_var).pack(fill="x", pady=1)

    Label(main_frame, text="Data inicial (dd/mm/aaaa):").pack(anchor="w")
    Entry(main_frame, textvariable=start_date_var).pack(fill="x", pady=1)

    Label(main_frame, text="Data final (dd/mm/aaaa):").pack(anchor="w")
    Entry(main_frame, textvariable=end_date_var).pack(fill="x", pady=1)

    Button(
        main_frame,
        text="Gerar relatório",
        command=lambda: execute_query(
            companies_var.get(),
            start_date_var.get(),
            end_date_var.get(),
            root
        )
    ).pack(pady=3)

    root.mainloop()


if __name__ == "__main__":
    window()
