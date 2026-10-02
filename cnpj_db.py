"""Persistência local (SQLite) do histórico de consultas e fretes.

Guarda só o necessário para o caso de uso (minimização de dados - LGPD art. 6º):
CNPJ, razão social, situação cadastral e dados de endereço/frete. Nada de
CPF de sócio ou outro dado pessoal que a API eventualmente retorne.
"""
import os
import sqlite3
from datetime import datetime, timedelta

DB_PATH = os.path.join(os.path.dirname(__file__), "data", "historico.db")


def _conectar():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_db():
    with _conectar() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS historico_consultas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cnpj TEXT NOT NULL,
                nome TEXT,
                situacao TEXT,
                municipio TEXT,
                uf TEXT,
                cep TEXT,
                data_hora TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS historico_fretes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cnpj TEXT NOT NULL,
                uf_destino TEXT,
                peso_kg REAL,
                valor REAL,
                prazo_dias INTEGER,
                data_hora TEXT NOT NULL
            )
        """)


def salvar_consulta(cnpj: str, nome: str, situacao: str, municipio: str, uf: str, cep: str = ""):
    with _conectar() as conn:
        conn.execute(
            "INSERT INTO historico_consultas (cnpj, nome, situacao, municipio, uf, cep, data_hora) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (cnpj, nome, situacao, municipio, uf, cep, datetime.now().isoformat(timespec="seconds")),
        )


def salvar_frete(cnpj: str, uf_destino: str, peso_kg: float, valor: float, prazo_dias: int):
    with _conectar() as conn:
        conn.execute(
            "INSERT INTO historico_fretes (cnpj, uf_destino, peso_kg, valor, prazo_dias, data_hora) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            (cnpj, uf_destino, peso_kg, valor, prazo_dias, datetime.now().isoformat(timespec="seconds")),
        )


def listar_consultas(limite: int = 200):
    with _conectar() as conn:
        cur = conn.execute(
            "SELECT * FROM historico_consultas ORDER BY data_hora DESC LIMIT ?", (limite,)
        )
        return [dict(r) for r in cur.fetchall()]


def listar_fretes(limite: int = 200):
    with _conectar() as conn:
        cur = conn.execute(
            "SELECT * FROM historico_fretes ORDER BY data_hora DESC LIMIT ?", (limite,)
        )
        return [dict(r) for r in cur.fetchall()]


def excluir_por_cnpj(cnpj: str):
    """Direito de exclusão (LGPD art. 18) - remove todo o histórico de um CNPJ."""
    with _conectar() as conn:
        conn.execute("DELETE FROM historico_consultas WHERE cnpj = ?", (cnpj,))
        conn.execute("DELETE FROM historico_fretes WHERE cnpj = ?", (cnpj,))


def limpar_historico_antigo(dias: int = 365):
    """Retenção limitada (LGPD art. 15/16) - apaga registros mais antigos que `dias`."""
    limite = (datetime.now() - timedelta(days=dias)).isoformat(timespec="seconds")
    with _conectar() as conn:
        cur1 = conn.execute("DELETE FROM historico_consultas WHERE data_hora < ?", (limite,))
        cur2 = conn.execute("DELETE FROM historico_fretes WHERE data_hora < ?", (limite,))
        return cur1.rowcount + cur2.rowcount
