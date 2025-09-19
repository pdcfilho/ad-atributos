import argparse
import openpyxl
from ldap3 import Server, Connection, ALL, NTLM

def main():
    parser = argparse.ArgumentParser(description="Exportar usuários do AD para Excel")
    parser.add_argument("--server", required=True, help="Servidor LDAP (ex: ldap://dominio.local)")
    parser.add_argument("--user", required=True, help="Usuário (ex: DOMINIO\\usuario ou usuario@dominio.local)")
    parser.add_argument("--password", required=True, help="Senha do usuário")
    parser.add_argument("--base", required=True, help="Base DN (ex: DC=empresa,DC=com,DC=br)")
    parser.add_argument("--output", default="usuarios.xlsx", help="Arquivo de saída XLSX")
    args = parser.parse_args()

    # Conexão
    server = Server(args.server, get_info=ALL)
    conn = Connection(
        server,
        user=args.user,
        password=args.password,
        authentication=NTLM,
        auto_bind=True
    )

    # Filtro: apenas usuários ativos, sem computadores
    search_filter = "(&(objectClass=user)(!(objectClass=computer))(!(userAccountControl:1.2.840.113556.1.4.803:=2)))"
    conn.search(
        search_base=args.base,
        search_filter=search_filter,
        attributes=["sAMAccountName", "mobile", "title"]
    )

    # Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Usuarios"
    ws.append(["Login", "Mobile", "Title"])

    for entry in conn.entries:
        login = str(entry.sAMAccountName) if "sAMAccountName" in entry else ""
        mobile = str(entry.mobile) if "mobile" in entry else ""
        title = str(entry.title) if "title" in entry else ""
        ws.append([login, mobile, title])

    wb.save(args.output)
    print(f"Arquivo {args.output} gerado com sucesso!")

if __name__ == "__main__":
    main()
