"""RPA visual de lançamentos para o sistema demonstrativo Pão D'Ouro RH."""
import os
import logging
import queue
import re
import threading
import time
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

BASE_URL = 'http://127.0.0.1:8000'
DELAY_ENTRE_FUNCIONARIOS = 2.0  # segundos; altere aqui se desejar
COLUNAS = ('funcionario_id', 'periodo', 'horas_trabalhadas', 'he50', 'he100', 'faltas', 'atestados')
ALIASES = {
    'id': 'funcionario_id', 'id_trabalhador': 'funcionario_id', 'id_funcionario': 'funcionario_id',
    'horas': 'horas_trabalhadas', 'horas_extras_50': 'he50', 'horas_extras_100': 'he100',
    'he_50': 'he50', 'he_100': 'he100', 'mes': 'periodo',
}


def normalizar(texto):
    import unicodedata
    texto = unicodedata.normalize('NFKD', str(texto or ''))
    texto = ''.join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r'[^a-z0-9]+', '_', texto.lower()).strip('_')


def converter_horas(valor, linha, campo):
    """Converte HH:MM (exportação de ponto) para horas decimais."""
    from datetime import time, timedelta
    if valor is None or str(valor).strip() == '':
        return 0.0
    if isinstance(valor, timedelta):
        return valor.total_seconds() / 3600
    if isinstance(valor, time):
        return valor.hour + valor.minute / 60 + valor.second / 3600
    if isinstance(valor, str) and ':' in valor:
        partes = valor.strip().split(':')
        if len(partes) not in (2, 3):
            raise ValueError(f'Linha {linha}: {campo} inválido: {valor}')
        try:
            horas, minutos = int(partes[0]), int(partes[1])
            segundos = int(partes[2]) if len(partes) == 3 else 0
        except ValueError as e:
            raise ValueError(f'Linha {linha}: {campo} inválido: {valor}') from e
        if horas < 0 or not 0 <= minutos < 60 or not 0 <= segundos < 60:
            raise ValueError(f'Linha {linha}: {campo} inválido: {valor}')
        return horas + minutos / 60 + segundos / 3600
    try:
        return float(str(valor).replace(',', '.'))
    except (ValueError, TypeError) as e:
        raise ValueError(f'Linha {linha}: {campo} inválido: {valor}') from e


def ler_excel(caminho):
    from datetime import datetime as DataHora, date
    wb = load_workbook(caminho, read_only=True, data_only=True)
    try:
        ws = wb.active
        linhas = ws.iter_rows(values_only=True)
        cabecalho = next(linhas, None)
        if not cabecalho:
            raise ValueError('A planilha está vazia.')
        aliases = {
            **ALIASES, 'horas_extras': 'he50', 'horas_extra': 'he50',
            'funcionario': 'nome', 'data': 'data', 'atestado': 'atestados',
            'horas_trabalhadas': 'horas_trabalhadas', 'he_50': 'he50',
            'he_100': 'he100', 'faltas': 'faltas',
        }
        nomes = [aliases.get(normalizar(c), normalizar(c)) for c in cabecalho]
        if 'funcionario_id' not in nomes or 'horas_trabalhadas' not in nomes:
            raise ValueError('Excel precisa ter ID e Horas trabalhadas.')
        if 'periodo' not in nomes and 'data' not in nomes:
            raise ValueError('Excel precisa ter Período ou Data.')
        # O relatório de ponto tem várias linhas por funcionário: consolidar por mês.
        agrupados = {}
        for numero, valores in enumerate(linhas, start=2):
            if not any(v is not None and str(v).strip() for v in valores):
                continue
            linha = dict(zip(nomes, valores))
            id_ = linha.get('funcionario_id')
            if id_ is None or not str(id_).strip():
                raise ValueError(f'Linha {numero}: ID vazio')
            if isinstance(id_, float) and id_.is_integer():
                id_ = int(id_)
            id_ = str(id_).strip()
            periodo = str(linha.get('periodo') or '').strip()
            if not periodo:
                data = linha.get('data')
                if isinstance(data, (DataHora, date)):
                    d = data
                else:
                    try:
                        d = DataHora.strptime(str(data).strip().split(' ')[0], '%d/%m/%Y')
                    except ValueError as e:
                        raise ValueError(f'Linha {numero}: Data inválida: {data}') from e
                meses = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                         'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro']
                periodo = f'{meses[d.month - 1]} de {d.year}'
            chave = (id_, periodo)
            if chave not in agrupados:
                agrupados[chave] = dict(funcionario_id=id_, periodo=periodo,
                    horas_trabalhadas=0.0, he50=0.0, he100=0.0,
                    faltas=0, atestados=0, _linha=numero)
            reg = agrupados[chave]
            for campo in ('horas_trabalhadas', 'he50', 'he100', 'faltas', 'atestados'):
                valor = converter_horas(linha.get(campo, 0), numero, campo)
                if valor < 0 or valor > 744:
                    raise ValueError(f'Linha {numero}: {campo} fora do intervalo 0–744')
                if campo in ('faltas', 'atestados') and not valor.is_integer():
                    raise ValueError(f'Linha {numero}: {campo} deve ser inteiro')
                reg[campo] += int(valor) if campo in ('faltas', 'atestados') else valor
        resultado = list(agrupados.values())
        if not resultado:
            raise ValueError('Nenhum funcionário encontrado.')
        for reg in resultado:
            for campo in ('horas_trabalhadas', 'he50', 'he100', 'faltas', 'atestados'):
                if reg[campo] > 744:
                    raise ValueError(f'Funcionário {reg["funcionario_id"]}: {campo} acima de 744 no mês')
                if campo not in ('faltas', 'atestados'):
                    reg[campo] = round(reg[campo], 2)
        return resultado
    finally:
        wb.close()


class Interface:
    def __init__(self, root):
        self.root = root
        self.root.title('RPA · Pão D’Ouro RH')
        self.root.geometry('800x570')
        self.fila = queue.Queue()
        self.pasta_logs = Path(__file__).resolve().parent / 'logs'
        self.pasta_logs.mkdir(parents=True, exist_ok=True)
        self.logger = logging.getLogger(f'RPA_RH_{id(self)}')
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        self.arquivo_log = None

        self.caminho = tk.StringVar()
        self.status = tk.StringVar(value='Aguardando planilha')
        corpo = ttk.Frame(root, padding=20)
        corpo.pack(fill='both', expand=True)
        ttk.Label(corpo, text='Robô de lançamentos de folha', font=('Segoe UI', 18, 'bold')).pack(anchor='w', pady=(0, 8))
        ttk.Label(corpo, text='Selecione o Excel, confira os dados e inicie a automação no navegador.').pack(anchor='w')
        linha = ttk.Frame(corpo)
        linha.pack(fill='x', pady=15)
        ttk.Entry(linha, textvariable=self.caminho).pack(side='left', fill='x', expand=True)
        ttk.Button(linha, text='Escolher Excel', command=self.escolher).pack(side='left', padx=(8, 0))
        self.botao = ttk.Button(corpo, text='Iniciar lançamentos', command=self.iniciar)
        self.botao.pack(anchor='w', pady=(0, 10))
        ttk.Button(corpo, text='Abrir pasta de logs', command=self.abrir_logs).pack(anchor='w', pady=(0, 10))
        ttk.Label(corpo, textvariable=self.status).pack(anchor='w', pady=(0, 10))
        self.progresso = ttk.Progressbar(corpo, mode='determinate')
        self.progresso.pack(fill='x', pady=(0, 15))
        self.log_texto = tk.Text(corpo, height=17, state='disabled', wrap='word', bg='#111827', fg='#e5e7eb', font=('Consolas', 10))
        self.log_texto.pack(fill='both', expand=True)
        self.root.after(100, self.atualizar)

    def escolher(self):
        arquivo = filedialog.askopenfilename(filetypes=[('Planilhas Excel', '*.xlsx')])
        if arquivo:
            self.caminho.set(arquivo)
            self.registrar('Excel selecionado: ' + Path(arquivo).name)

    def abrir_logs(self):
        if os.name == 'nt':
            os.startfile(str(self.pasta_logs))
        else:
            messagebox.showinfo('Pasta de logs', str(self.pasta_logs))

    def registrar(self, mensagem):
        self.logger.info(mensagem)
        self.fila.put(('log', mensagem))

    def iniciar(self):
        arquivo = Path(self.caminho.get())
        if not arquivo.is_file() or arquivo.suffix.lower() != '.xlsx':
            messagebox.showerror('Arquivo inválido', 'Selecione um arquivo .xlsx existente.')
            return
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)
            handler.close()
        self.arquivo_log = self.pasta_logs / ('execucao_' + datetime.now().strftime('%Y%m%d_%H%M%S_%f') + '.log')
        handler = logging.FileHandler(self.arquivo_log, encoding='utf-8')
        handler.setFormatter(logging.Formatter('[%(asctime)s] %(levelname)s - %(message)s', datefmt='%d/%m/%Y %H:%M:%S'))
        self.logger.addHandler(handler)
        self.registrar('Arquivo de log criado: ' + str(self.arquivo_log))
        self.botao.configure(state='disabled')
        self.progresso['value'] = 0
        threading.Thread(target=self.executar, args=(arquivo,), daemon=True).start()

    def atualizar(self):
        try:
            while True:
                tipo, valor = self.fila.get_nowait()
                if tipo == 'log':
                    self.log_texto.configure(state='normal')
                    self.log_texto.insert('end', f'[{datetime.now():%H:%M:%S}] {valor}\n')
                    self.log_texto.see('end')
                    self.log_texto.configure(state='disabled')
                    self.status.set(valor)
                elif tipo == 'progresso':
                    atual, total = valor
                    self.progresso['value'] = atual / total * 100
                elif tipo == 'fim':
                    self.botao.configure(state='normal')
                    if valor:
                        messagebox.showinfo('Finalizado', valor)
                    else:
                        messagebox.showwarning('Verificar logs', 'Execução encerrada com pendências. Consulte os logs.')
        except queue.Empty:
            pass
        self.root.after(100, self.atualizar)

    def executar(self, arquivo):
        sucesso = 0
        falhas = []
        duplicados = 0
        try:
            registros = ler_excel(arquivo)
            self.registrar(f'Excel recebido com sucesso: {len(registros)} funcionários/períodos consolidados')
            self.registrar('ATENÇÃO: horas extras sem classificação serão tratadas como HE 50%; HE 100% e atestados ausentes serão zero (somente demonstração).')
            periodos = {r['periodo'] for r in registros}
            if len(periodos) != 1:
                raise ValueError('Use um único período por execução para gerar o relatório correto.')
            periodo = next(iter(periodos))
            pasta = Path.home() / 'Downloads' / 'RPA_RH'
            pasta.mkdir(parents=True, exist_ok=True)
            with sync_playwright() as p:
                navegador = p.chromium.launch(headless=False)
                contexto = navegador.new_context(accept_downloads=True)
                pagina = contexto.new_page()
                self.registrar('Procedimento iniciado; abrindo sistema de RH')
                pagina.goto(BASE_URL + '/index.html', wait_until='domcontentloaded', timeout=15000)
                pagina.get_by_role('link', name='Processamento da folha').click()
                pagina.locator('#formFolha').wait_for()
                self.registrar('Página folha.html aberta com sucesso')
                self.registrar(f'Intervalo entre funcionários: {DELAY_ENTRE_FUNCIONARIOS:g} segundos')
                for indice, registro in enumerate(registros, start=1):
                    id_ = registro['funcionario_id']
                    try:
                        pagina.locator('#periodo').select_option(label=periodo)
                        pagina.locator('#idTrabalhador').fill(id_)
                        with pagina.expect_response(lambda r: '/api/funcionario?' in r.url, timeout=15000) as consulta:
                            pagina.locator('#buscarFuncionario').click()
                        resposta = consulta.value
                        if not resposta.ok:
                            raise ValueError(f'ID não encontrado ({resposta.status})')
                        nome = pagina.locator('#nome').input_value()
                        if not nome:
                            raise ValueError('Nome não preenchido após busca')
                        for campo in ('horas_trabalhadas', 'he50', 'he100', 'faltas', 'atestados'):
                            seletor = 'horasTrabalhadas' if campo == 'horas_trabalhadas' else campo
                            pagina.locator('#' + seletor).fill(str(registro[campo]))
                        # O sistema exibe alert() após cada lançamento.
                        alertas = []
                        def fechar_alerta(dialog):
                            alertas.append(dialog.message)
                            dialog.accept()
                        pagina.once('dialog', fechar_alerta)
                        with pagina.expect_response(lambda r: r.url.endswith('/api/lancamentos') and r.request.method == 'POST', timeout=15000) as lancamento:
                            pagina.locator('#formFolha button[type=submit]').click()
                        retorno = lancamento.value
                        if not retorno.ok:
                            try:
                                motivo = retorno.json().get('erro', str(retorno.status))
                            except Exception:
                                motivo = str(retorno.status)
                            raise ValueError(motivo)
                        sucesso += 1
                        self.registrar(f'[{indice}/{len(registros)}] Funcionário {id_} - {nome}: registrado com sucesso')
                    except Exception as erro:
                        mensagem = str(erro)
                        if 'Já existe lançamento para este funcionário e período' in mensagem or 'Ja existe lancamento para este funcionario e periodo' in mensagem:
                            duplicados += 1
                            self.registrar(f'DUPLICADO funcionário {id_}: já cadastrado no período; ignorando')
                        else:
                            falhas.append((id_, mensagem))
                            self.registrar(f'ERRO funcionário {id_} (linha {registro["_linha"]}): {mensagem}')
                    self.fila.put(('progresso', (indice, len(registros))))
                    if indice < len(registros):
                        time.sleep(DELAY_ENTRE_FUNCIONARIOS)
                self.registrar(f'Lançamentos concluídos: {sucesso} novos, {duplicados} duplicados ignorados, {len(falhas)} falhas')
                self.registrar('Abrindo relatórios e tentando baixar mesmo com duplicados ou falhas')
                pagina.get_by_role('link', name='Relatórios').click()
                pagina.locator('#periodoRelatorio').select_option(label=periodo)
                with pagina.expect_download(timeout=120000) as download_info:
                    pagina.locator('#baixarSalarios').click()
                download = download_info.value
                destino = pasta / download.suggested_filename
                if destino.exists():
                    destino = pasta / (destino.stem + '_' + datetime.now().strftime('%Y%m%d_%H%M%S') + destino.suffix)
                download.save_as(str(destino))
                self.registrar('Relatório baixado; abrindo arquivo: ' + str(destino))
                if os.name == 'nt':
                    os.startfile(str(destino))
                else:
                    self.registrar('Abra o arquivo salvo para visualizar o resultado.')
                resumo = f'{sucesso} novos, {duplicados} duplicados, {len(falhas)} falhas. Relatório: {destino}'
                self.registrar('Finalizado com pendências' if falhas else 'Finalizado com sucesso')
                self.fila.put(('fim', resumo))
                navegador.close()
        except Exception as erro:
            self.registrar('ERRO na execução: ' + str(erro))
            self.fila.put(('fim', None))


if __name__ == '__main__':
    root = tk.Tk()
    Interface(root)
    root.mainloop()
