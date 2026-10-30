"""Sistema de Controle de Estoque e Produção.

Aplicação desktop em Python (Tkinter + SQLite) para:
  * cadastrar, editar e excluir produtos;
  * registrar entradas, saídas e produção (com histórico);
  * alertar sobre produtos com estoque baixo;
  * exportar o estoque para CSV (abre direto no Excel).

Como executar:
    python estoque.py

Os dados ficam salvos no arquivo "estoque.db", na mesma pasta do programa.
"""

import csv
import sqlite3
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

ARQUIVO_BANCO = Path(__file__).with_name("estoque.db")
TIPOS_MOVIMENTO = ("Entrada", "Saída", "Produção")


# --------------------------------------------------------------------------
# Validação de dados
# --------------------------------------------------------------------------
class ErroValidacao(Exception):
    """Dado inválido informado pelo usuário (a mensagem é exibida na tela)."""


def validar_texto(valor, campo, obrigatorio=True):
    valor = valor.strip()
    if obrigatorio and not valor:
        raise ErroValidacao(f"O campo '{campo}' é obrigatório.")
    return valor


def validar_inteiro(valor, campo, minimo=0):
    try:
        numero = int(valor.strip())
    except ValueError:
        raise ErroValidacao(f"O campo '{campo}' deve ser um número inteiro.") from None
    if numero < minimo:
        raise ErroValidacao(f"O campo '{campo}' deve ser maior ou igual a {minimo}.")
    return numero


def validar_preco(valor, campo="Preço"):
    try:
        numero = float(valor.strip().replace(",", "."))
    except ValueError:
        raise ErroValidacao(f"O campo '{campo}' deve ser um número (ex.: 12,50).") from None
    if numero < 0:
        raise ErroValidacao(f"O campo '{campo}' não pode ser negativo.")
    return round(numero, 2)


def formatar_moeda(valor):
    texto = f"{valor:,.2f}"
    return "R$ " + texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_data(texto_iso):
    return datetime.fromisoformat(texto_iso).strftime("%d/%m/%Y %H:%M")


# --------------------------------------------------------------------------
# Camada de dados (SQLite)
# --------------------------------------------------------------------------
class Banco:
    """Acesso ao banco de dados. Não conhece nada sobre a interface gráfica."""

    def __init__(self, caminho=ARQUIVO_BANCO):
        self.conn = sqlite3.connect(caminho)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._criar_tabelas()

    def _criar_tabelas(self):
        with self.conn:
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS produtos (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    codigo     TEXT    NOT NULL UNIQUE,
                    nome       TEXT    NOT NULL,
                    categoria  TEXT    NOT NULL DEFAULT '',
                    quantidade INTEGER NOT NULL DEFAULT 0 CHECK (quantidade >= 0),
                    minimo     INTEGER NOT NULL DEFAULT 0 CHECK (minimo >= 0),
                    preco      REAL    NOT NULL DEFAULT 0 CHECK (preco >= 0)
                )
                """
            )
            self.conn.execute(
                """
                CREATE TABLE IF NOT EXISTS movimentos (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    produto_id INTEGER NOT NULL
                               REFERENCES produtos(id) ON DELETE CASCADE,
                    tipo       TEXT    NOT NULL,
                    quantidade INTEGER NOT NULL CHECK (quantidade > 0),
                    data       TEXT    NOT NULL,
                    observacao TEXT    NOT NULL DEFAULT ''
                )
                """
            )

    # ---- produtos --------------------------------------------------------
    def listar_produtos(self, busca=""):
        termo = f"%{busca.strip()}%"
        return self.conn.execute(
            "SELECT * FROM produtos WHERE codigo LIKE ? OR nome LIKE ? "
            "OR categoria LIKE ? ORDER BY nome COLLATE NOCASE",
            (termo, termo, termo),
        ).fetchall()

    def buscar_produto(self, produto_id):
        return self.conn.execute(
            "SELECT * FROM produtos WHERE id = ?", (produto_id,)
        ).fetchone()

    def listar_estoque_baixo(self):
        return self.conn.execute(
            "SELECT * FROM produtos WHERE quantidade <= minimo "
            "ORDER BY quantidade, nome COLLATE NOCASE"
        ).fetchall()

    def adicionar_produto(self, codigo, nome, categoria, quantidade, minimo, preco):
        try:
            with self.conn:
                self.conn.execute(
                    "INSERT INTO produtos (codigo, nome, categoria, quantidade, minimo, preco) "
                    "VALUES (?, ?, ?, ?, ?, ?)",
                    (codigo, nome, categoria, quantidade, minimo, preco),
                )
        except sqlite3.IntegrityError:
            raise ErroValidacao(f"Já existe um produto com o código '{codigo}'.") from None

    def atualizar_produto(self, produto_id, codigo, nome, categoria, minimo, preco, **_):
        """Atualiza os dados cadastrais. A quantidade só muda por movimentação."""
        try:
            with self.conn:
                self.conn.execute(
                    "UPDATE produtos SET codigo = ?, nome = ?, categoria = ?, "
                    "minimo = ?, preco = ? WHERE id = ?",
                    (codigo, nome, categoria, minimo, preco, produto_id),
                )
        except sqlite3.IntegrityError:
            raise ErroValidacao(f"Já existe um produto com o código '{codigo}'.") from None

    def excluir_produto(self, produto_id):
        with self.conn:
            self.conn.execute("DELETE FROM produtos WHERE id = ?", (produto_id,))

    # ---- movimentações ---------------------------------------------------
    def registrar_movimento(self, produto_id, tipo, quantidade, observacao=""):
        """Registra a movimentação e atualiza o estoque na mesma transação."""
        if tipo not in TIPOS_MOVIMENTO:
            raise ErroValidacao("Tipo de movimentação inválido.")
        produto = self.buscar_produto(produto_id)
        if produto is None:
            raise ErroValidacao("Produto não encontrado.")

        if tipo == "Saída":
            if quantidade > produto["quantidade"]:
                raise ErroValidacao(
                    f"Estoque insuficiente: há apenas {produto['quantidade']} "
                    f"unidade(s) de '{produto['nome']}'."
                )
            delta = -quantidade
        else:  # Entrada e Produção aumentam o estoque
            delta = quantidade

        with self.conn:
            self.conn.execute(
                "UPDATE produtos SET quantidade = quantidade + ? WHERE id = ?",
                (delta, produto_id),
            )
            self.conn.execute(
                "INSERT INTO movimentos (produto_id, tipo, quantidade, data, observacao) "
                "VALUES (?, ?, ?, ?, ?)",
                (produto_id, tipo, quantidade, datetime.now().isoformat(timespec="seconds"), observacao),
            )
        return produto["quantidade"] + delta

    def listar_movimentos(self, limite=300):
        return self.conn.execute(
            "SELECT m.data, p.codigo, p.nome, m.tipo, m.quantidade, m.observacao "
            "FROM movimentos m JOIN produtos p ON p.id = m.produto_id "
            "ORDER BY m.id DESC LIMIT ?",
            (limite,),
        ).fetchall()


# --------------------------------------------------------------------------
# Interface gráfica (Tkinter)
# --------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self, banco):
        super().__init__()
        self.banco = banco
        self.produto_selecionado = None   # id do produto carregado no formulário
        self.mapa_produtos = {}           # texto do combobox -> id do produto

        self.title("Controle de Estoque e Produção")
        self.geometry("1000x640")
        self.minsize(900, 560)
        ttk.Style(self).theme_use("clam")

        abas = ttk.Notebook(self)
        abas.pack(fill="both", expand=True, padx=8, pady=8)
        self.aba_produtos = ttk.Frame(abas)
        self.aba_movimentos = ttk.Frame(abas)
        self.aba_alertas = ttk.Frame(abas)
        abas.add(self.aba_produtos, text="  Produtos  ")
        abas.add(self.aba_movimentos, text="  Movimentações  ")
        abas.add(self.aba_alertas, text="  Estoque baixo  ")

        self._montar_aba_produtos()
        self._montar_aba_movimentos()
        self._montar_aba_alertas()
        self.atualizar_tudo()

    # ---- utilitários de layout ------------------------------------------
    @staticmethod
    def _criar_tabela(pai, colunas, altura=12):
        """Cria um Treeview com barra de rolagem. colunas = [(id, titulo, largura, alinhamento)]."""
        quadro = ttk.Frame(pai)
        tabela = ttk.Treeview(
            quadro, columns=[c[0] for c in colunas], show="headings", height=altura
        )
        for ident, titulo, largura, alinhamento in colunas:
            tabela.heading(ident, text=titulo)
            tabela.column(ident, width=largura, anchor=alinhamento)
        barra = ttk.Scrollbar(quadro, orient="vertical", command=tabela.yview)
        tabela.configure(yscrollcommand=barra.set)
        tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")
        tabela.tag_configure("baixo", background="#fde2e2")
        return quadro, tabela

    # ---- aba Produtos -----------------------------------------------------
    def _montar_aba_produtos(self):
        form = ttk.LabelFrame(self.aba_produtos, text="Dados do produto", padding=10)
        form.pack(fill="x", padx=10, pady=(10, 5))

        rotulos = [
            ("codigo", "Código"), ("nome", "Nome"), ("categoria", "Categoria"),
            ("quantidade", "Quantidade inicial"), ("minimo", "Estoque mínimo"),
            ("preco", "Preço (R$)"),
        ]
        self.campos = {}
        for i, (chave, rotulo) in enumerate(rotulos):
            linha, coluna = divmod(i, 3)
            ttk.Label(form, text=rotulo).grid(row=linha * 2, column=coluna, sticky="w", padx=6)
            entrada = ttk.Entry(form, width=30)
            entrada.grid(row=linha * 2 + 1, column=coluna, sticky="w", padx=6, pady=(0, 8))
            self.campos[chave] = entrada

        botoes = ttk.Frame(form)
        botoes.grid(row=4, column=0, columnspan=3, sticky="w", padx=6)
        self.btn_cadastrar = ttk.Button(botoes, text="Cadastrar", command=self.cadastrar)
        self.btn_atualizar = ttk.Button(botoes, text="Atualizar", command=self.atualizar_produto)
        self.btn_excluir = ttk.Button(botoes, text="Excluir", command=self.excluir)
        self.btn_limpar = ttk.Button(botoes, text="Limpar", command=self.limpar_formulario)
        for botao in (self.btn_cadastrar, self.btn_atualizar, self.btn_excluir, self.btn_limpar):
            botao.pack(side="left", padx=(0, 6))

        barra_busca = ttk.Frame(self.aba_produtos)
        barra_busca.pack(fill="x", padx=10, pady=(8, 2))
        ttk.Label(barra_busca, text="Buscar:").pack(side="left")
        self.var_busca = tk.StringVar()
        self.var_busca.trace_add("write", lambda *_: self.atualizar_tabela_produtos())
        ttk.Entry(barra_busca, textvariable=self.var_busca, width=35).pack(side="left", padx=6)
        ttk.Button(barra_busca, text="Exportar CSV", command=self.exportar_csv).pack(side="right")

        colunas = [
            ("codigo", "Código", 100, "w"), ("nome", "Nome", 280, "w"),
            ("categoria", "Categoria", 150, "w"), ("quantidade", "Qtd.", 70, "center"),
            ("minimo", "Mínimo", 70, "center"), ("preco", "Preço", 110, "e"),
        ]
        quadro, self.tabela_produtos = self._criar_tabela(self.aba_produtos, colunas, altura=12)
        quadro.pack(fill="both", expand=True, padx=10, pady=(2, 10))
        self.tabela_produtos.bind("<<TreeviewSelect>>", self.carregar_no_formulario)
        self._atualizar_botoes()

    def _ler_formulario(self):
        """Lê e valida os campos. Lança ErroValidacao se algo estiver errado."""
        c = self.campos
        return {
            "codigo": validar_texto(c["codigo"].get(), "Código").upper(),
            "nome": validar_texto(c["nome"].get(), "Nome"),
            "categoria": validar_texto(c["categoria"].get(), "Categoria", obrigatorio=False),
            "quantidade": validar_inteiro(c["quantidade"].get(), "Quantidade inicial"),
            "minimo": validar_inteiro(c["minimo"].get(), "Estoque mínimo"),
            "preco": validar_preco(c["preco"].get()),
        }

    def _atualizar_botoes(self):
        editando = self.produto_selecionado is not None
        self.btn_cadastrar.state(["disabled"] if editando else ["!disabled"])
        self.btn_atualizar.state(["!disabled"] if editando else ["disabled"])
        self.btn_excluir.state(["!disabled"] if editando else ["disabled"])
        # A quantidade só é editável no cadastro; depois, apenas por movimentação.
        self.campos["quantidade"].state(["disabled"] if editando else ["!disabled"])

    def cadastrar(self):
        try:
            self.banco.adicionar_produto(**self._ler_formulario())
        except ErroValidacao as erro:
            messagebox.showwarning("Dado inválido", str(erro))
            return
        self.limpar_formulario()
        self.atualizar_tudo()

    def atualizar_produto(self):
        try:
            dados = self._ler_formulario()
            self.banco.atualizar_produto(self.produto_selecionado, **dados)
        except ErroValidacao as erro:
            messagebox.showwarning("Dado inválido", str(erro))
            return
        self.limpar_formulario()
        self.atualizar_tudo()

    def excluir(self):
        produto = self.banco.buscar_produto(self.produto_selecionado)
        if not messagebox.askyesno(
            "Confirmar exclusão",
            f"Excluir '{produto['nome']}' e todo o seu histórico de movimentações?",
        ):
            return
        self.banco.excluir_produto(self.produto_selecionado)
        self.limpar_formulario()
        self.atualizar_tudo()

    def limpar_formulario(self):
        self.produto_selecionado = None
        self.tabela_produtos.selection_remove(self.tabela_produtos.selection())
        self._atualizar_botoes()  # reabilita o campo quantidade antes de limpar
        for entrada in self.campos.values():
            entrada.delete(0, tk.END)

    def carregar_no_formulario(self, _evento=None):
        selecao = self.tabela_produtos.selection()
        if not selecao:
            return
        produto = self.banco.buscar_produto(int(selecao[0]))
        self.produto_selecionado = produto["id"]
        self.campos["quantidade"].state(["!disabled"])
        for chave, entrada in self.campos.items():
            valor = produto[chave]
            if chave == "preco":
                valor = f"{valor:.2f}".replace(".", ",")
            entrada.delete(0, tk.END)
            entrada.insert(0, valor)
        self._atualizar_botoes()

    def atualizar_tabela_produtos(self):
        tabela = self.tabela_produtos
        tabela.delete(*tabela.get_children())
        for p in self.banco.listar_produtos(self.var_busca.get()):
            tag = ("baixo",) if p["quantidade"] <= p["minimo"] else ()
            tabela.insert(
                "", "end", iid=str(p["id"]), tags=tag,
                values=(p["codigo"], p["nome"], p["categoria"], p["quantidade"],
                        p["minimo"], formatar_moeda(p["preco"])),
            )

    def exportar_csv(self):
        caminho = filedialog.asksaveasfilename(
            defaultextension=".csv", initialfile="estoque.csv",
            filetypes=[("Arquivo CSV", "*.csv")],
        )
        if not caminho:
            return
        # utf-8-sig + ";" para abrir corretamente no Excel em português
        with open(caminho, "w", newline="", encoding="utf-8-sig") as arquivo:
            escritor = csv.writer(arquivo, delimiter=";")
            escritor.writerow(["Código", "Nome", "Categoria", "Quantidade", "Mínimo", "Preço"])
            for p in self.banco.listar_produtos():
                escritor.writerow([p["codigo"], p["nome"], p["categoria"], p["quantidade"],
                                   p["minimo"], f"{p['preco']:.2f}".replace(".", ",")])
        messagebox.showinfo("Exportação concluída", f"Arquivo salvo em:\n{caminho}")

    # ---- aba Movimentações ------------------------------------------------
    def _montar_aba_movimentos(self):
        form = ttk.LabelFrame(self.aba_movimentos, text="Nova movimentação", padding=10)
        form.pack(fill="x", padx=10, pady=10)

        ttk.Label(form, text="Produto").grid(row=0, column=0, sticky="w", padx=6)
        self.combo_produto = ttk.Combobox(form, width=42, state="readonly")
        self.combo_produto.grid(row=1, column=0, padx=6, pady=(0, 8), sticky="w")

        ttk.Label(form, text="Tipo").grid(row=0, column=1, sticky="w", padx=6)
        self.combo_tipo = ttk.Combobox(form, width=14, state="readonly", values=TIPOS_MOVIMENTO)
        self.combo_tipo.current(0)
        self.combo_tipo.grid(row=1, column=1, padx=6, pady=(0, 8), sticky="w")

        ttk.Label(form, text="Quantidade").grid(row=0, column=2, sticky="w", padx=6)
        self.entrada_qtd = ttk.Entry(form, width=12)
        self.entrada_qtd.grid(row=1, column=2, padx=6, pady=(0, 8), sticky="w")

        ttk.Label(form, text="Observação").grid(row=2, column=0, sticky="w", padx=6)
        self.entrada_obs = ttk.Entry(form, width=45)
        self.entrada_obs.grid(row=3, column=0, columnspan=2, padx=6, sticky="w")
        ttk.Button(form, text="Registrar", command=self.registrar_movimento).grid(
            row=3, column=2, padx=6, sticky="w"
        )

        colunas = [
            ("data", "Data", 130, "w"), ("codigo", "Código", 90, "w"),
            ("nome", "Produto", 240, "w"), ("tipo", "Tipo", 90, "center"),
            ("quantidade", "Qtd.", 70, "center"), ("observacao", "Observação", 280, "w"),
        ]
        quadro, self.tabela_movimentos = self._criar_tabela(self.aba_movimentos, colunas, altura=14)
        quadro.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def registrar_movimento(self):
        try:
            texto = self.combo_produto.get()
            if texto not in self.mapa_produtos:
                raise ErroValidacao("Selecione um produto da lista.")
            produto_id = self.mapa_produtos[texto]
            quantidade = validar_inteiro(self.entrada_qtd.get(), "Quantidade", minimo=1)
            novo_saldo = self.banco.registrar_movimento(
                produto_id, self.combo_tipo.get(), quantidade, self.entrada_obs.get().strip()
            )
        except ErroValidacao as erro:
            messagebox.showwarning("Não foi possível registrar", str(erro))
            return

        self.entrada_qtd.delete(0, tk.END)
        self.entrada_obs.delete(0, tk.END)
        self.atualizar_tudo()

        produto = self.banco.buscar_produto(produto_id)
        if novo_saldo <= produto["minimo"]:
            messagebox.showwarning(
                "Estoque baixo",
                f"'{produto['nome']}' está com {novo_saldo} unidade(s) "
                f"(mínimo: {produto['minimo']}).",
            )

    def atualizar_tabela_movimentos(self):
        tabela = self.tabela_movimentos
        tabela.delete(*tabela.get_children())
        for m in self.banco.listar_movimentos():
            tabela.insert("", "end", values=(
                formatar_data(m["data"]), m["codigo"], m["nome"],
                m["tipo"], m["quantidade"], m["observacao"],
            ))

    # ---- aba Estoque baixo ------------------------------------------------
    def _montar_aba_alertas(self):
        self.rotulo_alertas = ttk.Label(self.aba_alertas, font=("TkDefaultFont", 11, "bold"))
        self.rotulo_alertas.pack(anchor="w", padx=12, pady=(12, 6))
        colunas = [
            ("codigo", "Código", 100, "w"), ("nome", "Nome", 300, "w"),
            ("quantidade", "Qtd. atual", 90, "center"), ("minimo", "Mínimo", 90, "center"),
            ("falta", "Repor (sugestão)", 130, "center"),
        ]
        quadro, self.tabela_alertas = self._criar_tabela(self.aba_alertas, colunas, altura=16)
        quadro.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def atualizar_tabela_alertas(self):
        tabela = self.tabela_alertas
        tabela.delete(*tabela.get_children())
        produtos = self.banco.listar_estoque_baixo()
        for p in produtos:
            tabela.insert("", "end", tags=("baixo",), values=(
                p["codigo"], p["nome"], p["quantidade"], p["minimo"],
                max(p["minimo"] * 2 - p["quantidade"], 1),
            ))
        self.rotulo_alertas.config(
            text=f"{len(produtos)} produto(s) com estoque igual ou abaixo do mínimo"
        )

    # ---- atualização geral ------------------------------------------------
    def atualizar_tudo(self):
        self.atualizar_tabela_produtos()
        self.atualizar_tabela_movimentos()
        self.atualizar_tabela_alertas()
        produtos = self.banco.listar_produtos()
        self.mapa_produtos = {f"{p['codigo']} - {p['nome']}": p["id"] for p in produtos}
        self.combo_produto["values"] = list(self.mapa_produtos)
        if self.combo_produto.get() not in self.mapa_produtos:
            self.combo_produto.set("")

        # Mantém o produto em edição sincronizado (ex.: depois de uma movimentação)
        if self.produto_selecionado is not None:
            if self.tabela_produtos.exists(str(self.produto_selecionado)):
                self.tabela_produtos.selection_set(str(self.produto_selecionado))
            else:
                self.limpar_formulario()


def main():
    App(Banco()).mainloop()


if __name__ == "__main__":
    main()
