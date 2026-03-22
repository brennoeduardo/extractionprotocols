"""
Extrator de Protocolos Financeiros
===================================
Regras de negócio:
  - Processa apenas UM nível de subpastas (as "pastas protocolo")
  - Dentro de cada pasta protocolo:
      .pdf            → move para a pasta de destino
      .txt            → converte para PDF (via ReportLab) e move
      .xls / .xlsx    → ignorado com aviso (Excel não é esperado)
      subpasta        → ignorada com aviso (sinal de algo errado)
      qualquer outro  → ignorado com aviso

Dependência extra: pip install reportlab
"""

import os
import re
import queue
import shutil
import threading
import logging
from datetime import datetime
from pathlib import Path

import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph
    REPORTLAB_OK = True
except ImportError:
    REPORTLAB_OK = False

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ────────────────────────────────────────────────────────────────────────────
#  Extensões aceitas / ignoradas
# ────────────────────────────────────────────────────────────────────────────

EXT_MOVER     = {".pdf"}
EXT_CONVERTER = {".txt"}
EXT_EXCEL     = {".xls", ".xlsx", ".xlsm", ".xlsb"}

# O nome do arquivo (sem extensão) deve conter o padrão PROJETO(9)-NOTA(9)-PROTOCOLO(14)
# e ter apenas dígitos e hífens (sem letras). Exemplos válidos:
#   010204358-000003766-00006302460000
#   20260329-010203963-000009521-00006470780000
PADRAO_NOME = re.compile(r"^[\d-]*\d{9}-\d{9}-\d{14}[\d-]*$")


# ────────────────────────────────────────────────────────────────────────────
#  Utilitários
# ────────────────────────────────────────────────────────────────────────────

def configurar_logger(caminho_log: str) -> logging.Logger:
    logger = logging.getLogger("extrator")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()
    handler = logging.FileHandler(caminho_log, encoding="utf-8")
    handler.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(handler)
    return logger


def gerar_nome_unico(destino: str) -> str:
    if not os.path.exists(destino):
        return destino
    base, ext = os.path.splitext(destino)
    i = 1
    while os.path.exists(f"{base}_{i}{ext}"):
        i += 1
    return f"{base}_{i}{ext}"


# Detecta prefixo de data no formato AAAAMMDD seguido de hífen
# Exemplo: 20260329-010203963-... -> 2026-03-29-010203963-...
_PREFIXO_DATA = re.compile(r"^(\d{4})(\d{2})(\d{2})-(.+)$")

def formatar_nome(nome_sem_ext: str) -> tuple:
    """
    Retorna (nome_final, foi_renomeado).
    Se o nome começa com uma data AAAAMMDD, formata para AAAA-MM-DD.
    Caso contrário, devolve o nome sem alteração.
    """
    m = _PREFIXO_DATA.match(nome_sem_ext)
    if m:
        ano, mes, dia, resto = m.groups()
        return f"{ano}-{mes}-{dia}-{resto}", True
    return nome_sem_ext, False


def txt_para_pdf(caminho_txt: str, caminho_pdf: str) -> None:
    """
    Converte um arquivo .txt para .pdf preservando o conteúdo como texto.
    Usa ReportLab com fonte monospace para manter indentação.
    Lança exceção se ReportLab não estiver instalado ou a conversão falhar.
    """
    if not REPORTLAB_OK:
        raise ImportError(
            "ReportLab não está instalado. Execute: pip install reportlab"
        )

    with open(caminho_txt, encoding="utf-8", errors="replace") as f:
        linhas = f.readlines()

    styles = getSampleStyleSheet()
    estilo = styles["Normal"]
    estilo.fontName = "Courier"
    estilo.fontSize = 9
    estilo.leading = 12

    doc = SimpleDocTemplate(
        caminho_pdf,
        pagesize=A4,
        leftMargin=2 * cm,
        rightMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
    )

    elementos = []
    for linha in linhas:
        texto = linha.rstrip("\n").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        elementos.append(Paragraph(texto if texto else "&nbsp;", estilo))

    doc.build(elementos)


def contar_arquivos_protocolo(pasta_raiz: str) -> int:
    """Conta apenas arquivos diretos dentro das subpastas de primeiro nível."""
    total = 0
    try:
        for nome in os.listdir(pasta_raiz):
            subpasta = os.path.join(pasta_raiz, nome)
            if os.path.isdir(subpasta):
                total += sum(
                    1 for item in os.listdir(subpasta)
                    if os.path.isfile(os.path.join(subpasta, item))
                )
    except PermissionError:
        pass
    return total


# ────────────────────────────────────────────────────────────────────────────
#  Worker (roda em thread separada)
# ────────────────────────────────────────────────────────────────────────────

def processar_worker(
    pasta_raiz: str,
    pasta_destino: str,
    fila_ui: queue.Queue,
    logger: logging.Logger,
) -> None:

    contadores = {
        "movidos": 0,
        "convertidos": 0,
        "erros": 0,
        "excel": 0,
        "outros": 0,
        "fora_padrao": 0,
        "subpastas": 0,
        "vazias": 0,
    }

    def log(msg: str, nivel: str = "info") -> None:
        getattr(logger, nivel)(msg)
        fila_ui.put(("log", msg))

    try:
        subpastas = sorted([
            os.path.join(pasta_raiz, nome)
            for nome in os.listdir(pasta_raiz)
            if os.path.isdir(os.path.join(pasta_raiz, nome))
        ])

        if not subpastas:
            log("Nenhuma subpasta encontrada na pasta raiz.")
            return

        total_itens = sum(len(os.listdir(p)) for p in subpastas)
        fila_ui.put(("total", max(total_itens, 1)))
        itens_processados = 0

        os.makedirs(pasta_destino, exist_ok=True)

        for pasta in subpastas:
            nome_pasta = os.path.basename(pasta)
            itens = sorted(os.listdir(pasta))

            if not itens:
                log(f"AVISO  | Pasta vazia → [{nome_pasta}]")
                contadores["vazias"] += 1
                continue

            for item in itens:
                caminho_item = os.path.join(pasta, item)

                # ── Caso 0: Nome não segue o padrão → ignorar ─────────────
                nome_sem_ext = Path(item).stem
                if not os.path.isdir(caminho_item) and not PADRAO_NOME.match(nome_sem_ext):
                    log(
                        f"AVISO  | Nome fora do padrão em [{nome_pasta}] → [{item}] "
                    )
                    contadores["fora_padrao"] += 1
                    itens_processados += 1
                    fila_ui.put(("progresso", itens_processados / total_itens))
                    continue

                # ── Caso 1: É uma subpasta → ignorar ──────────────────────
                if os.path.isdir(caminho_item):
                    log(
                        f"AVISO  | Subpasta ignorada em [{nome_pasta}] → [{item}] "
                        f"(protocolos não deveriam ter subpastas)"
                    )
                    contadores["subpastas"] += 1
                    itens_processados += 1
                    fila_ui.put(("progresso", itens_processados / total_itens))
                    continue

                ext = Path(item).suffix.lower()

                # ── Caso 2: PDF → mover (renomeando data se necessário) ───
                if ext in EXT_MOVER:
                    nome_formatado, tinha_data = formatar_nome(nome_sem_ext)
                    nome_destino = nome_formatado + ext
                    destino = gerar_nome_unico(os.path.join(pasta_destino, nome_destino))
                    try:
                        shutil.move(caminho_item, destino)
                        nome_final = os.path.basename(destino)
                        if tinha_data:
                            log(f"OK     | PDF movido [{item}] → [{nome_final}] (data formatada) de [{nome_pasta}]")
                        elif nome_final != item:
                            log(f"OK     | PDF movido [{item}] → [{nome_final}] de [{nome_pasta}]")
                        else:
                            log(f"OK     | PDF movido [{item}] de [{nome_pasta}]")
                        contadores["movidos"] += 1
                    except PermissionError:
                        log(f"ERRO   | Sem permissão para mover [{item}] em [{nome_pasta}]", "error")
                        contadores["erros"] += 1
                    except OSError as e:
                        log(f"ERRO   | Falha ao mover [{item}] — {e}", "error")
                        contadores["erros"] += 1

                # ── Caso 3: TXT → converter para PDF e mover ──────────────
                elif ext in EXT_CONVERTER:
                    nome_formatado, tinha_data = formatar_nome(nome_sem_ext)
                    nome_pdf = nome_formatado + ".pdf"
                    pdf_temp = os.path.join(pasta, nome_pdf)
                    destino  = gerar_nome_unico(os.path.join(pasta_destino, nome_pdf))
                    try:
                        txt_para_pdf(caminho_item, pdf_temp)
                        shutil.move(pdf_temp, destino)
                        os.remove(caminho_item)
                        nome_final = os.path.basename(destino)
                        if tinha_data:
                            log(f"OK     | TXT convertido [{item}] → [{nome_final}] (data formatada) de [{nome_pasta}]")
                        else:
                            log(f"OK     | TXT convertido [{item}] → [{nome_final}] de [{nome_pasta}]")
                        contadores["convertidos"] += 1
                    except ImportError as e:
                        log(f"ERRO   | {e}", "error")
                        contadores["erros"] += 1
                    except Exception as e:
                        log(f"ERRO   | Falha ao converter TXT [{item}] — {e}", "error")
                        if os.path.exists(pdf_temp):
                            try:
                                os.remove(pdf_temp)
                            except OSError:
                                pass
                        contadores["erros"] += 1

                # ── Caso 4: Excel → ignorar com aviso ─────────────────────
                elif ext in EXT_EXCEL:
                    log(
                        f"AVISO  | Excel ignorado em [{nome_pasta}] → [{item}] "
                        f"(somente PDF e TXT são processados)"
                    )
                    contadores["excel"] += 1

                # ── Caso 5: Qualquer outro tipo → ignorar ──────────────────
                else:
                    log(
                        f"AVISO  | Tipo ignorado [{ext or 'sem extensão'}] em "
                        f"[{nome_pasta}] → [{item}]"
                    )
                    contadores["outros"] += 1

                itens_processados += 1
                fila_ui.put(("progresso", itens_processados / total_itens))

    except Exception as e:
        log(f"ERRO CRÍTICO: {e}", "critical")

    finally:
        log("─── RESUMO ──────────────────────────────────────")
        log(f"PDFs movidos            : {contadores['movidos']}")
        log(f"TXTs convertidos→PDF    : {contadores['convertidos']}")
        log(f"Erros                   : {contadores['erros']}")
        log(f"Fora do padrão de nome  : {contadores['fora_padrao']}")
        log(f"Excels ignorados        : {contadores['excel']}")
        log(f"Outros tipos ignorados  : {contadores['outros']}")
        log(f"Subpastas ignoradas     : {contadores['subpastas']}")
        log(f"Pastas protocolo vazias : {contadores['vazias']}")
        fila_ui.put(("concluido", contadores))


# ────────────────────────────────────────────────────────────────────────────
#  Interface gráfica
# ────────────────────────────────────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Extrator de Protocolos Financeiros")
        self.geometry("740x580")
        self.resizable(False, False)

        self.pasta_raiz    = ""
        self.pasta_destino = ""
        self.fila_ui: queue.Queue = queue.Queue()
        self.logger: logging.Logger | None = None
        self._total_itens = 1
        self._processados = 0

        self._aviso_reportlab()
        self._construir_ui()

    # ── Aviso de dependência ─────────────────────────────────────────────

    def _aviso_reportlab(self) -> None:
        if not REPORTLAB_OK:
            messagebox.showwarning(
                "Dependência ausente",
                "ReportLab não está instalado.\n\n"
                "A conversão de TXT → PDF ficará desabilitada.\n\n"
                "Para ativar, execute no terminal:\n"
                "    pip install reportlab",
            )

    # ── Construção da UI ─────────────────────────────────────────────────

    def _construir_ui(self) -> None:
        pad = {"padx": 20}

        # ── Origem ──────────────────────────────────────────────────────
        frame_origem = ctk.CTkFrame(self)
        frame_origem.pack(fill="x", **pad, pady=(14, 4))

        ctk.CTkLabel(frame_origem, text="Pasta de origem (protocolos):").pack(
            anchor="w", padx=10, pady=(8, 2)
        )
        frame_origem_row = ctk.CTkFrame(frame_origem, fg_color="transparent")
        frame_origem_row.pack(fill="x", padx=10, pady=(0, 8))

        self.entry_origem = ctk.CTkEntry(frame_origem_row, height=34)
        self.entry_origem.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            frame_origem_row, text="Selecionar Pasta",
            width=150, height=34,
            command=self._selecionar_origem,
        ).pack(side="left")

        # ── Destino ──────────────────────────────────────────────────────
        frame_destino = ctk.CTkFrame(self)
        frame_destino.pack(fill="x", **pad, pady=4)

        ctk.CTkLabel(frame_destino, text="Pasta de destino dos arquivos:").pack(
            anchor="w", padx=10, pady=(8, 2)
        )
        frame_destino_row = ctk.CTkFrame(frame_destino, fg_color="transparent")
        frame_destino_row.pack(fill="x", padx=10, pady=(0, 8))

        self.entry_destino = ctk.CTkEntry(frame_destino_row, height=34)
        self.entry_destino.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            frame_destino_row, text="Selecionar Pasta",
            width=150, height=34,
            command=self._selecionar_destino,
        ).pack(side="left")

        # Progresso
        frame_prog = ctk.CTkFrame(self)
        frame_prog.pack(fill="x", **pad, pady=4)

        self.label_status = ctk.CTkLabel(frame_prog, text="Aguardando…")
        self.label_status.pack(anchor="w", padx=8)

        self.progress = ctk.CTkProgressBar(frame_prog, width=680)
        self.progress.set(0)
        self.progress.pack(padx=8, pady=4)

        self.label_contador = ctk.CTkLabel(frame_prog, text="")
        self.label_contador.pack(anchor="e", padx=8)

        # Log
        self.log_area = ScrolledText(
            self, height=13, width=88,
            bg="#1a1a1a", fg="#d4d4d4",
            font=("Consolas", 9),
        )
        self.log_area.pack(**pad, pady=4)

        # Botão
        self.btn_process = ctk.CTkButton(
            self, text="Iniciar Processamento",
            command=self._confirmar_e_processar, height=38,
        )
        self.btn_process.pack(pady=8)

    # ── Seleção de pastas ────────────────────────────────────────────────

    def _selecionar_origem(self) -> None:
        pasta = filedialog.askdirectory(title="Pasta com as subpastas de protocolo")
        if pasta:
            self.pasta_raiz = pasta
            self.entry_origem.delete(0, "end")
            self.entry_origem.insert(0, pasta)
            if not self.pasta_destino:
                self.pasta_destino = pasta
                self.entry_destino.delete(0, "end")
                self.entry_destino.insert(0, pasta)

    def _selecionar_destino(self) -> None:
        pasta = filedialog.askdirectory(title="Pasta de destino dos arquivos extraídos")
        if pasta:
            self.pasta_destino = pasta
            self.entry_destino.delete(0, "end")
            self.entry_destino.insert(0, pasta)

    # ── Confirmação e início ─────────────────────────────────────────────

    def _confirmar_e_processar(self) -> None:
        self.pasta_raiz    = self.entry_origem.get().strip()
        self.pasta_destino = self.entry_destino.get().strip()

        if not self.pasta_raiz or not os.path.isdir(self.pasta_raiz):
            messagebox.showerror("Erro", "Selecione uma pasta de origem válida.")
            return
        if not self.pasta_destino:
            messagebox.showerror("Erro", "Selecione uma pasta de destino.")
            return

        total = contar_arquivos_protocolo(self.pasta_raiz)
        if total == 0:
            messagebox.showwarning(
                "Aviso",
                "Nenhum arquivo encontrado nas pastas protocolo.\n"
                "Verifique se a pasta de origem contém subpastas com documentos.",
            )
            return

        confirmado = messagebox.askyesno(
            "Confirmar processamento",
            f"Serão processados {total} arquivo(s) encontrados nas subpastas de:\n\n"
            f"  {self.pasta_raiz}\n\n"
            f"Destino: {self.pasta_destino}\n\n"
            "Esta operação NÃO pode ser desfeita. Continuar?",
        )
        if not confirmado:
            return

        self._iniciar_processamento(total)

    def _iniciar_processamento(self, total: int) -> None:
        caminho_log = os.path.join(self.pasta_raiz, "log_processamento.txt")
        self.logger = configurar_logger(caminho_log)

        self.log_area.delete("1.0", "end")
        self.progress.set(0)
        self.label_status.configure(text="Processando…")
        self.label_contador.configure(text=f"0 / {total}")
        self.btn_process.configure(state="disabled", text="Processando…")

        self._total_itens = max(total, 1)
        self._processados = 0

        thread = threading.Thread(
            target=processar_worker,
            args=(self.pasta_raiz, self.pasta_destino, self.fila_ui, self.logger),
            daemon=True,
        )
        thread.start()
        self.after(100, self._verificar_fila)

    # ── Polling da fila UI ───────────────────────────────────────────────

    def _verificar_fila(self) -> None:
        try:
            while True:
                tipo, dado = self.fila_ui.get_nowait()

                if tipo == "log":
                    self._append_log(dado)

                elif tipo == "total":
                    self._total_itens = dado
                    self.label_contador.configure(text=f"0 / {dado}")

                elif tipo == "progresso":
                    self._processados = int(dado * self._total_itens)
                    self.progress.set(dado)
                    self.label_contador.configure(
                        text=f"{self._processados} / {self._total_itens}"
                    )

                elif tipo == "concluido":
                    self._on_concluido(dado)
                    return

        except queue.Empty:
            pass

        self.after(100, self._verificar_fila)

    def _append_log(self, mensagem: str) -> None:
        hora = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert("end", f"[{hora}] {mensagem}\n")
        self.log_area.see("end")

    def _on_concluido(self, c: dict) -> None:
        self.progress.set(1)
        self.label_status.configure(text="Concluído.")
        self.label_contador.configure(
            text=f"{self._total_itens} / {self._total_itens}"
        )
        self.btn_process.configure(state="normal", text="Iniciar Processamento")

        if self.logger:
            for h in self.logger.handlers:
                h.close()
            self.logger.handlers.clear()

        avisos = []
        if c["fora_padrao"]:
            avisos.append(f"⚠ {c['fora_padrao']} arquivo(s) com nome fora do padrão ignorado(s)")
        if c["subpastas"]:
            avisos.append(f"⚠ {c['subpastas']} subpasta(s) ignorada(s) nos protocolos")
        if c["excel"]:
            avisos.append(f"⚠ {c['excel']} Excel(s) ignorado(s)")
        if c["outros"]:
            avisos.append(f"⚠ {c['outros']} arquivo(s) de outro tipo ignorado(s)")
        if c["vazias"]:
            avisos.append(f"⚠ {c['vazias']} pasta(s) protocolo vazia(s)")

        corpo = (
            f"✔ {c['movidos']} PDF(s) movido(s)\n"
            f"✖ {c['erros']} erro(s)\n"
        )
        if avisos:
            corpo += "\n" + "\n".join(avisos)
        corpo += "\n\nConsulte log_processamento.txt para detalhes."

        titulo = "Concluído com erros" if c["erros"] else "Concluído com sucesso"
        if c["erros"]:
            messagebox.showwarning(titulo, corpo)
        else:
            messagebox.showinfo(titulo, corpo)


# ────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()