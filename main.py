import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pythoncom
import win32com.client

# Pasta onde estão os arquivos do CorelDRAW
pasta_cdr_padrao = r"\\neoserver\HD NEOCOOLER\CLASSE TURISTA\NEOCOOLER\-----  CIDADES  -----\ARRAIAL DO CABO\COREL DRAW\BOLSA PRAIA"

# Pasta onde será salvo o TIFF
pasta_tif_padrao = r"\\neoserver\HD NEOCOOLER\TesteNeocooler"

COR_FUNDO = "#0b0f14"
COR_PAINEL = "#111923"
COR_CIANO = "#00f5ff"
COR_VERDE = "#39ff14"
COR_ROSA = "#ff2bd6"
COR_TEXTO = "#e6f1ff"

CDR_FILTER_TIFF = 772
CDR_EXPORT_CURRENT_PAGE = 1
CDR_COMPRESSION_JPEG = 8
CDR_IMAGE_TYPE_POR_MODELO = {
    0: 4,  # RGB: 24 bits
    1: 5,  # CMYK: 32 bits
    2: 2,  # Escala de cinza: 8 bits
    3: 7,  # Lab: 24 bits
}


def caminho_recurso(nome):
    pasta_base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(pasta_base, nome)


def criar_opcoes_exportacao(corel, doc, modo_cor):
    if modo_cor == "RGB":
        tipo_imagem = CDR_IMAGE_TYPE_POR_MODELO[0]
    elif modo_cor == "CMYK":
        tipo_imagem = CDR_IMAGE_TYPE_POR_MODELO[1]
    elif modo_cor == "Original":
        modelo_cor = int(doc.ColorContext.BlendingColorModel)
        try:
            tipo_imagem = CDR_IMAGE_TYPE_POR_MODELO[modelo_cor]
        except KeyError:
            raise ValueError(
                f"Modelo de cores não suportado para exportação TIFF/JPEG: {modelo_cor}."
            ) from None
    else:
        raise ValueError(f"Modo de cor não suportado: {modo_cor}.")

    opcoes = corel.CreateStructExportOptions()
    opcoes.ImageType = tipo_imagem
    opcoes.UseColorProfile = True
    opcoes.Compression = CDR_COMPRESSION_JPEG
    return opcoes


def criar_caminhos_exportacao(
    pasta_destino, nome_base, numeros_paginas, sufixos_paginas
):
    caminhos = []
    caminhos_reservados = set()

    for numero_pagina in numeros_paginas:
        if numero_pagina == 1:
            sufixo = ""
        elif numero_pagina <= len(sufixos_paginas) + 1:
            sufixo = sufixos_paginas[numero_pagina - 2]
            if not sufixo.strip():
                sufixo = ""
        elif sufixos_paginas and sufixos_paginas[-1].strip():
            sufixo = f"{sufixos_paginas[-1]}{numero_pagina}"
        else:
            sufixo = ""

        nome_arquivo = f"{nome_base}_Pagina{numero_pagina}{sufixo}"
        caminho = os.path.join(pasta_destino, f"{nome_arquivo}.tif")
        chave_caminho = os.path.normcase(os.path.abspath(caminho))

        if chave_caminho in caminhos_reservados or os.path.exists(caminho):
            nome_arquivo = f"{nome_arquivo} - página {numero_pagina}"
            caminho = os.path.join(pasta_destino, f"{nome_arquivo}.tif")
            chave_caminho = os.path.normcase(os.path.abspath(caminho))

        numero_versao = 2
        while os.path.exists(caminho) or chave_caminho in caminhos_reservados:
            caminho = os.path.join(
                pasta_destino, f"{nome_arquivo} ({numero_versao}).tif"
            )
            chave_caminho = os.path.normcase(os.path.abspath(caminho))
            numero_versao += 1

        caminhos.append(caminho)
        caminhos_reservados.add(chave_caminho)

    return caminhos


def converter_pastas(
    pasta_cdr,
    pasta_tif,
    registrar_log,
    cancelamento,
    modo_cor,
    sufixos_paginas,
):
    inicio = time.perf_counter()
    quantidade_exportada = 0
    erros = []
    cancelada = False

    corel = None

    try:
        pastas_e_arquivos = []
        quantidade_encontrada = 0
        for pasta_atual, subpastas, arquivos in os.walk(pasta_cdr):
            if cancelamento.is_set():
                cancelada = True
                break
            subpastas.sort()
            arquivos_cdr = sorted(
                arquivo for arquivo in arquivos if arquivo.lower().endswith(".cdr")
            )
            quantidade_encontrada += len(arquivos_cdr)
            pastas_e_arquivos.append((pasta_atual, arquivos_cdr))

        registrar_log(f"{quantidade_encontrada} arquivo(s) CDR encontrado(s).")
        if cancelada:
            registrar_log("Exportação cancelada antes de iniciar.")
            return quantidade_exportada, time.perf_counter() - inicio, erros, True
        if not quantidade_encontrada:
            registrar_log("Nenhum arquivo .cdr encontrado para exportar.")
            return quantidade_exportada, time.perf_counter() - inicio, erros, False

        corel = win32com.client.DispatchEx("CorelDRAW.Application")
        corel.Visible = False

        for pasta_atual, arquivos in pastas_e_arquivos:
            if cancelamento.is_set():
                cancelada = True
                break
            pasta_relativa = os.path.relpath(pasta_atual, pasta_cdr)
            pasta_tif_atual = (
                pasta_tif
                if pasta_relativa == os.curdir
                else os.path.join(pasta_tif, pasta_relativa)
            )
            os.makedirs(pasta_tif_atual, exist_ok=True)

            for nome_arquivo in arquivos:
                if cancelamento.is_set():
                    cancelada = True
                    break

                arquivo_cdr = os.path.join(pasta_atual, nome_arquivo)
                nome_sem_extensao = os.path.splitext(nome_arquivo)[0]

                doc = None
                try:
                    registrar_log(f"Exportando: {arquivo_cdr}")
                    doc = corel.OpenDocument(arquivo_cdr)
                    doc.Activate()
                    quantidade_paginas = int(doc.Pages.Count)
                    if quantidade_paginas < 1:
                        raise ValueError(
                            f"O documento não contém páginas: {arquivo_cdr}"
                        )
                    registrar_log(
                        f"{quantidade_paginas} página(s) encontrada(s) em "
                        f"{nome_arquivo}."
                    )
                    quantidade_paginas_exportar = min(
                        quantidade_paginas, len(sufixos_paginas) + 1
                    )
                    numeros_paginas = range(1, quantidade_paginas_exportar + 1)
                    if quantidade_paginas_exportar < quantidade_paginas:
                        registrar_log(
                            f"{quantidade_paginas - quantidade_paginas_exportar} "
                            f"página(s) ignorada(s): quantidade de inputs configurados "
                            f"para {nome_arquivo}."
                        )
                    paginas = [
                        (numero_pagina, doc.Pages.Item(numero_pagina))
                        for numero_pagina in numeros_paginas
                    ]
                    opcoes = criar_opcoes_exportacao(corel, doc, modo_cor)
                    opcoes_paleta = corel.CreateStructPaletteOptions()
                    caminhos_exportacao = criar_caminhos_exportacao(
                        pasta_tif_atual,
                        nome_sem_extensao,
                        numeros_paginas,
                        sufixos_paginas,
                    )
                    for (numero_pagina, pagina), arquivo_tif in zip(
                        paginas, caminhos_exportacao
                    ):
                        if cancelamento.is_set():
                            cancelada = True
                            break
                        try:
                            registrar_log(
                                f"Exportando página {numero_pagina} de "
                                f"{quantidade_paginas}: {arquivo_tif}"
                            )
                            pagina.Activate()
                            filtro = doc.ExportEx(
                                arquivo_tif,
                                CDR_FILTER_TIFF,
                                CDR_EXPORT_CURRENT_PAGE,
                                opcoes,
                                opcoes_paleta,
                            )
                            if filtro is None:
                                raise RuntimeError(
                                    "O CorelDRAW não iniciou a exportação TIFF."
                                )
                            filtro.Finish()
                            quantidade_exportada += 1
                            registrar_log(
                                f"Concluído (página {numero_pagina}): {arquivo_tif}"
                            )
                        except Exception as erro:
                            mensagem = (
                                f"{arquivo_cdr}, página {numero_pagina}: {erro}"
                            )
                            erros.append(mensagem)
                            registrar_log(f"ERRO ao exportar {mensagem}", "erro")
                            print(f"Erro ao exportar {mensagem}", file=sys.stderr)
                except Exception as erro:
                    mensagem = f"{arquivo_cdr}: {erro}"
                    erros.append(mensagem)
                    registrar_log(f"ERRO ao exportar {mensagem}", "erro")
                    print(f"Erro ao exportar {mensagem}", file=sys.stderr)
                finally:
                    if doc is not None:
                        try:
                            doc.Close()
                        except Exception as erro:
                            mensagem = f"Falha ao fechar {arquivo_cdr}: {erro}"
                            erros.append(mensagem)
                            registrar_log(mensagem, "erro")
                            print(mensagem, file=sys.stderr)
            if cancelada:
                break
        if cancelada:
            registrar_log(
                "Cancelamento concluído. A página que estava em andamento foi finalizada."
            )
    except Exception as erro:
        mensagem = f"Falha geral na conversão: {erro}"
        erros.append(mensagem)
        registrar_log(mensagem, "erro")
        print(mensagem, file=sys.stderr)
    finally:
        if corel is not None:
            try:
                corel.Quit()
            except Exception as erro:
                mensagem = f"Falha ao encerrar o CorelDRAW: {erro}"
                erros.append(mensagem)
                registrar_log(mensagem, "erro")
                print(mensagem, file=sys.stderr)

    return quantidade_exportada, time.perf_counter() - inicio, erros, cancelada


def selecionar_pasta(campo):
    pasta = filedialog.askdirectory(
        title="Selecione uma pasta",
        initialdir=campo.get() or os.path.expanduser("~"),
    )
    if pasta:
        campo.set(pasta)


def iniciar_exportacao(
    janela,
    campo_cdr,
    campo_tif,
    campo_modo_cor,
    campos_sufixo_pagina,
    botao,
    botao_parar,
    botoes_pasta,
    campos_opcao,
    status,
    progresso,
    log,
):
    pasta_cdr = campo_cdr.get().strip()
    pasta_tif = campo_tif.get().strip()
    modo_cor = campo_modo_cor.get()
    sufixos_paginas = [campo.get() for campo in campos_sufixo_pagina]

    if not os.path.isdir(pasta_cdr):
        messagebox.showerror("Pasta inválida", "Selecione uma pasta CDR válida.")
        return
    if not pasta_tif:
        messagebox.showerror("Pasta inválida", "Selecione uma pasta para os TIFFs.")
        return

    botao.config(state=tk.DISABLED)
    botao_parar.config(state=tk.NORMAL)
    for botao_pasta in botoes_pasta:
        botao_pasta.config(state=tk.DISABLED)
    for campo_opcao, _estado in campos_opcao:
        campo_opcao.config(state=tk.DISABLED)
    log.delete("1.0", tk.END)
    status.set("Aguarde um pouquinho... conversão em andamento.")
    progresso.start(12)
    eventos = queue.Queue()
    cancelamento = threading.Event()

    def executar_conversao():
        com_inicializado = False
        try:
            pythoncom.CoInitialize()
            com_inicializado = True
            resultado = converter_pastas(
                pasta_cdr,
                pasta_tif,
                lambda mensagem, tipo="info": eventos.put(("log", mensagem, tipo)),
                cancelamento,
                modo_cor,
                sufixos_paginas,
            )
            eventos.put(("concluido", resultado))
        except Exception as erro:
            eventos.put(("falha", erro))
        finally:
            if com_inicializado:
                pythoncom.CoUninitialize()

    def atualizar_interface():
        finalizado = False
        try:
            while True:
                evento = eventos.get_nowait()
                if evento[0] == "log":
                    _, mensagem, tipo = evento
                    log.insert(tk.END, mensagem + "\n", tipo)
                    log.see(tk.END)
                elif evento[0] == "concluido":
                    _, (quantidade, duracao, erros, cancelada) = evento
                    progresso.stop()
                    if cancelada:
                        status.set(
                            f"Exportação parada: {quantidade} arquivo(s) TIFF exportado(s) "
                            f"em {duracao:.2f} segundos."
                        )
                    else:
                        status.set(
                            f"Finalizado: {quantidade} arquivo(s) TIFF exportado(s) "
                            f"em {duracao:.2f} segundos."
                        )
                    botao.config(state=tk.NORMAL)
                    botao_parar.config(state=tk.DISABLED)
                    for botao_pasta in botoes_pasta:
                        botao_pasta.config(state=tk.NORMAL)
                    for campo_opcao, estado in campos_opcao:
                        campo_opcao.config(state=estado)
                    resumo = (
                        f"Arquivos TIFF exportados: {quantidade}\n"
                        f"Tempo decorrido: {duracao:.2f} segundos"
                    )
                    if erros:
                        resumo += f"\nFalhas: {len(erros)}"
                        messagebox.showwarning(
                            "Exportação parada com falhas" if cancelada else
                            "Exportação concluída com falhas",
                            resumo,
                        )
                    elif cancelada:
                        messagebox.showinfo("Exportação parada", resumo)
                    else:
                        messagebox.showinfo("Exportação concluída", resumo)
                    finalizado = True
                elif evento[0] == "falha":
                    _, erro = evento
                    progresso.stop()
                    status.set("A exportação não pôde ser concluída.")
                    botao.config(state=tk.NORMAL)
                    botao_parar.config(state=tk.DISABLED)
                    for botao_pasta in botoes_pasta:
                        botao_pasta.config(state=tk.NORMAL)
                    for campo_opcao, estado in campos_opcao:
                        campo_opcao.config(state=estado)
                    log.insert(tk.END, f"ERRO: {erro}\n", "erro")
                    log.see(tk.END)
                    messagebox.showerror(
                        "Erro durante a exportação",
                        f"{erro}\n\nVerifique os caminhos e se o CorelDRAW está disponível.",
                    )
                    finalizado = True
        except queue.Empty:
            pass

        if not finalizado:
            janela.after(100, atualizar_interface)

    def parar_exportacao():
        cancelamento.set()
        botao_parar.config(state=tk.DISABLED)
        status.set("Parada solicitada. Aguarde a página atual terminar...")
        log.insert(
            tk.END,
            "Parada solicitada; aguardando a página atual terminar.\n",
            "erro",
        )
        log.see(tk.END)

    botao_parar.config(command=parar_exportacao)
    threading.Thread(target=executar_conversao, daemon=True).start()
    janela.after(100, atualizar_interface)


def criar_interface():
    janela = tk.Tk()
    janela.title("NeoCorel | CDR para TIFF")
    janela.iconbitmap(caminho_recurso("neocoreal.ico"))
    janela.configure(bg=COR_FUNDO)
    janela.resizable(True, True)
    janela.minsize(760, 560)

    campo_cdr = tk.StringVar(value=pasta_cdr_padrao)
    campo_tif = tk.StringVar(value=pasta_tif_padrao)
    campo_modo_cor = tk.StringVar(value="Original")
    campos_sufixo_pagina = []
    entradas_sufixo_pagina = []
    etiquetas_sufixo_pagina = []
    botoes_remover_sufixo = []

    janela.grid_columnconfigure(0, weight=1)

    tk.Label(
        janela,
        text="NEOCOREAL  /  EXPORTAÇÃO TIFF",
        bg=COR_FUNDO,
        fg=COR_CIANO,
        font=("Segoe UI", 16, "bold"),
    ).grid(row=0, column=0, columnspan=2, padx=16, pady=(16, 12), sticky="w")

    tk.Label(
        janela, text="Pasta dos arquivos CDR:", bg=COR_FUNDO, fg=COR_TEXTO
    ).grid(row=1, column=0, padx=16, pady=(4, 4), sticky="w")
    entrada_cdr = tk.Entry(
        janela,
        textvariable=campo_cdr,
        bg=COR_PAINEL,
        fg=COR_TEXTO,
        insertbackground=COR_CIANO,
        relief=tk.FLAT,
        highlightthickness=1,
        highlightbackground=COR_CIANO,
        highlightcolor=COR_CIANO,
    )
    entrada_cdr.grid(row=2, column=0, padx=(16, 6), pady=4, sticky="ew")
    botao_cdr = tk.Button(
        janela,
        text="Selecionar...",
        command=lambda: selecionar_pasta(campo_cdr),
        bg=COR_PAINEL,
        fg=COR_CIANO,
        activebackground=COR_CIANO,
        activeforeground=COR_FUNDO,
        relief=tk.FLAT,
        padx=12,
    )
    botao_cdr.grid(row=2, column=1, padx=(6, 16), pady=4)

    tk.Label(
        janela, text="Pasta de destino dos TIFFs:", bg=COR_FUNDO, fg=COR_TEXTO
    ).grid(row=3, column=0, padx=16, pady=(10, 4), sticky="w")
    entrada_tif = tk.Entry(
        janela,
        textvariable=campo_tif,
        bg=COR_PAINEL,
        fg=COR_TEXTO,
        insertbackground=COR_CIANO,
        relief=tk.FLAT,
        highlightthickness=1,
        highlightbackground=COR_CIANO,
        highlightcolor=COR_CIANO,
    )
    entrada_tif.grid(row=4, column=0, padx=(16, 6), pady=4, sticky="ew")
    botao_tif = tk.Button(
        janela,
        text="Selecionar...",
        command=lambda: selecionar_pasta(campo_tif),
        bg=COR_PAINEL,
        fg=COR_CIANO,
        activebackground=COR_CIANO,
        activeforeground=COR_FUNDO,
        relief=tk.FLAT,
        padx=12,
    )
    botao_tif.grid(row=4, column=1, padx=(6, 16), pady=4)

    opcoes_exportacao = tk.Frame(janela, bg=COR_FUNDO)
    opcoes_exportacao.grid(
        row=5, column=0, columnspan=2, padx=16, pady=(12, 4), sticky="ew"
    )
    tk.Label(
        opcoes_exportacao,
        text="Modo de cor:",
        bg=COR_FUNDO,
        fg=COR_TEXTO,
    ).grid(row=0, column=0, padx=(0, 6), sticky="w")
    seletor_modo_cor = ttk.Combobox(
        opcoes_exportacao,
        textvariable=campo_modo_cor,
        values=("Original", "RGB", "CMYK"),
        state="readonly",
        width=12,
    )
    seletor_modo_cor.grid(row=0, column=1, padx=(0, 18), sticky="w")
    quadro_sufixos = tk.Frame(opcoes_exportacao, bg=COR_FUNDO)
    quadro_sufixos.grid(row=1, column=0, columnspan=4, pady=(8, 0), sticky="ew")
    quadro_sufixos.grid_columnconfigure(0, weight=1)
    lista_sufixos = tk.Canvas(
        quadro_sufixos,
        height=92,
        bg=COR_FUNDO,
        highlightthickness=1,
        highlightbackground=COR_CIANO,
    )
    barra_sufixos = ttk.Scrollbar(
        quadro_sufixos, orient="vertical", command=lista_sufixos.yview
    )
    lista_sufixos.configure(yscrollcommand=barra_sufixos.set)
    lista_sufixos.grid(row=0, column=0, sticky="ew")
    barra_sufixos.grid(row=0, column=1, sticky="ns")
    quadro_campos_sufixo = tk.Frame(lista_sufixos, bg=COR_FUNDO)
    janela_campos_sufixo = lista_sufixos.create_window(
        (0, 0), window=quadro_campos_sufixo, anchor="nw"
    )

    def atualizar_tamanho_lista(_evento=None):
        lista_sufixos.configure(scrollregion=lista_sufixos.bbox("all"))

    def ajustar_largura_lista(evento):
        lista_sufixos.itemconfigure(janela_campos_sufixo, width=evento.width)

    quadro_campos_sufixo.bind("<Configure>", atualizar_tamanho_lista)
    lista_sufixos.bind("<Configure>", ajustar_largura_lista)

    def atualizar_linhas_sufixo():
        for indice, (etiqueta, entrada, botao_remover) in enumerate(
            zip(
                etiquetas_sufixo_pagina,
                entradas_sufixo_pagina,
                botoes_remover_sufixo,
            )
        ):
            linha = indice + 1
            etiqueta.config(text=f"Sufixo da página {indice + 2}:")
            etiqueta.grid(row=linha, column=0, padx=8, pady=4, sticky="w")
            entrada.grid(row=linha, column=1, padx=8, pady=4, sticky="ew")
            botao_remover.grid(row=linha, column=2, padx=8, pady=4)
        atualizar_tamanho_lista()

    def adicionar_campo_sufixo(valor=""):
        campo_sufixo = tk.StringVar(value=valor)
        campos_sufixo_pagina.append(campo_sufixo)
        etiqueta_sufixo = tk.Label(
            quadro_campos_sufixo,
            text="",
            bg=COR_FUNDO,
            fg=COR_TEXTO,
        )
        entrada_sufixo = tk.Entry(
            quadro_campos_sufixo,
            textvariable=campo_sufixo,
            bg=COR_PAINEL,
            fg=COR_TEXTO,
            insertbackground=COR_CIANO,
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=COR_CIANO,
            highlightcolor=COR_CIANO,
        )
        botao_remover = tk.Button(
            quadro_campos_sufixo,
            text="❌",
            command=lambda campo=campo_sufixo: remover_campo_sufixo(campo),
            bg=COR_PAINEL,
            fg=COR_ROSA,
            activebackground=COR_ROSA,
            activeforeground=COR_FUNDO,
            relief=tk.FLAT,
            padx=8,
        )
        etiquetas_sufixo_pagina.append(etiqueta_sufixo)
        entradas_sufixo_pagina.append(entrada_sufixo)
        botoes_remover_sufixo.append(botao_remover)
        atualizar_linhas_sufixo()
        lista_sufixos.after_idle(lambda: lista_sufixos.yview_moveto(1.0))

    def remover_campo_sufixo(campo_sufixo):
        indice = campos_sufixo_pagina.index(campo_sufixo)
        campos_sufixo_pagina.pop(indice)
        etiquetas_sufixo_pagina.pop(indice).destroy()
        entradas_sufixo_pagina.pop(indice).destroy()
        botoes_remover_sufixo.pop(indice).destroy()
        atualizar_linhas_sufixo()

    quadro_campos_sufixo.grid_columnconfigure(1, weight=1)
    adicionar_campo_sufixo(" FUNDO")
    adicionar_campo_sufixo(" TIRA")
    botao_adicionar_pagina = tk.Button(
        opcoes_exportacao,
        text="ADICIONAR PÁGINA",
        command=adicionar_campo_sufixo,
        bg=COR_PAINEL,
        fg=COR_CIANO,
        activebackground=COR_CIANO,
        activeforeground=COR_FUNDO,
        relief=tk.FLAT,
        padx=12,
    )
    botao_adicionar_pagina.grid(row=0, column=2, padx=(0, 8), sticky="w")
    opcoes_exportacao.grid_columnconfigure(3, weight=1)

    botoes_acao = tk.Frame(janela, bg=COR_FUNDO)
    botoes_acao.grid(row=6, column=0, columnspan=2, pady=(8, 10))
    botao_exportar = tk.Button(
        botoes_acao,
        text="CONVERTER PARA TIFF",
        bg=COR_VERDE,
        fg=COR_FUNDO,
        activebackground=COR_CIANO,
        activeforeground=COR_FUNDO,
        relief=tk.FLAT,
        font=("Segoe UI", 10, "bold"),
        padx=18,
        pady=8,
    )
    botao_exportar.grid(row=0, column=0, padx=6)
    botao_parar = tk.Button(
        botoes_acao,
        text="PARAR EXPORTAÇÃO",
        command=lambda: None,
        bg=COR_PAINEL,
        fg=COR_ROSA,
        activebackground=COR_ROSA,
        activeforeground=COR_FUNDO,
        relief=tk.FLAT,
        font=("Segoe UI", 10, "bold"),
        padx=18,
        pady=8,
        state=tk.DISABLED,
    )
    botao_parar.grid(row=0, column=1, padx=6)

    estilo = ttk.Style(janela)
    estilo.theme_use("clam")
    estilo.configure(
        "Neon.Horizontal.TProgressbar",
        troughcolor=COR_PAINEL,
        background=COR_CIANO,
        bordercolor=COR_PAINEL,
        lightcolor=COR_CIANO,
        darkcolor=COR_CIANO,
    )
    progresso = ttk.Progressbar(
        janela, mode="indeterminate", style="Neon.Horizontal.TProgressbar"
    )
    progresso.grid(row=7, column=0, columnspan=2, padx=16, sticky="ew")

    status = tk.StringVar(value="Selecione as pastas e inicie a conversão.")
    tk.Label(
        janela,
        textvariable=status,
        bg=COR_FUNDO,
        fg=COR_ROSA,
        font=("Segoe UI", 10, "bold"),
        anchor="w",
    ).grid(row=8, column=0, columnspan=2, padx=16, pady=(6, 6), sticky="ew")

    tk.Label(
        janela,
        text="LOG DA EXPORTAÇÃO",
        bg=COR_FUNDO,
        fg=COR_CIANO,
        font=("Segoe UI", 9, "bold"),
    ).grid(row=9, column=0, columnspan=2, padx=16, pady=(4, 4), sticky="w")
    moldura_log = tk.Frame(janela, bg=COR_CIANO, padx=1, pady=1)
    moldura_log.grid(
        row=10, column=0, columnspan=2, padx=16, pady=(0, 16), sticky="nsew"
    )
    janela.grid_rowconfigure(10, weight=1)
    log = tk.Text(
        moldura_log,
        height=12,
        bg="#070a0e",
        fg=COR_TEXTO,
        insertbackground=COR_CIANO,
        relief=tk.FLAT,
        wrap=tk.WORD,
        state=tk.NORMAL,
        font=("Consolas", 9),
    )
    barra_log = tk.Scrollbar(moldura_log, command=log.yview)
    log.configure(yscrollcommand=barra_log.set)
    log.tag_configure("info", foreground=COR_CIANO)
    log.tag_configure("erro", foreground=COR_ROSA)
    log.grid(row=0, column=0, sticky="nsew")
    barra_log.grid(row=0, column=1, sticky="ns")
    moldura_log.grid_columnconfigure(0, weight=1)
    moldura_log.grid_rowconfigure(0, weight=1)

    botao_exportar.config(
        command=lambda: iniciar_exportacao(
            janela,
            campo_cdr,
            campo_tif,
            campo_modo_cor,
            campos_sufixo_pagina,
            botao_exportar,
            botao_parar,
            (botao_cdr, botao_tif),
            tuple(
                [(seletor_modo_cor, "readonly"), (botao_adicionar_pagina, tk.NORMAL)]
                + [(entrada, tk.NORMAL) for entrada in entradas_sufixo_pagina]
                + [(botao, tk.NORMAL) for botao in botoes_remover_sufixo]
            ),
            status,
            progresso,
            log,
        )
    )

    janela.mainloop()


if __name__ == "__main__":
    try:
        criar_interface()
    except Exception as erro:
        print(f"Erro ao iniciar o programa: {erro}", file=sys.stderr)
        raise