import csv
import hashlib
import json
import re
import threading
import unicodedata
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter
from openpyxl import load_workbook

BASE = Path(__file__).resolve().parent
def normalize_name(value):
    text = unicodedata.normalize('NFKD', str(value))
    text = ''.join(c for c in text if not unicodedata.combining(c))
    return ' '.join(re.findall(r'[A-Z0-9]+', text.upper()))


def match_names(text, people):
    normalized = ' ' + normalize_name(text) + ' '
    return [name for name in people if ' ' + name + ' ' in normalized]


def split_previous_attempts(people, digest, previous):
    records = [r for r in previous if r.get('pdf') == digest]
    emails = {str(r.get('email', '')).strip().casefold() for r in records}
    eligible, blocked = [], []
    for key, person in people.items():
        # Missing recipient in a legacy record cannot be safely attributed.
        already_attempted = '' in emails or person['email'].strip().casefold() in emails
        (blocked if already_attempted else eligible).append(key)
    return eligible, blocked


def read_people(path, skipped=None):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        rows = iter(workbook.active.iter_rows(values_only=True))
        headers = [str(v or '').strip().lower().replace('-', '').replace('_', '') for v in next(rows)]
        aliases = {'nome': ['nome', 'funcionário', 'funcionario'], 'email': ['email', 'e mail', 'e-mail']}
        columns = {}
        for field, names in aliases.items():
            matches = [i for i, h in enumerate(headers) if h in names]
            if len(matches) != 1:
                raise ValueError(f'A primeira linha deve ter uma única coluna {field}.')
            columns[field] = matches[0]
        people = {}
        for line, row in enumerate(rows, 2):
            if not any(v is not None for v in row):
                continue
            values = {key: row[index] if index < len(row) else None for key, index in columns.items()}
            name, email = str(values['nome'] or '').strip(), str(values['email'] or '').strip()
            if not email:
                if skipped is not None:
                    skipped.append(f'Linha {line}: {name or "sem nome"} — sem e-mail')
                continue
            key = normalize_name(name)
            if len(key.split()) < 2 or not re.fullmatch(r'[^\s@;,<>]+@[^\s@;,<>]+\.[^\s@;,<>]+', email):
                raise ValueError(f'Linha {line}: confira o nome completo (ao menos duas palavras) e um único e-mail.')
            if key in people:
                raise ValueError(f'Linha {line}: Nome completo duplicado (inclusive após ignorar acentos). Não é possível distinguir homônimos pelo nome.')
            people[key] = {'nome': name, 'email': email, 'pages': []}
        if not people:
            raise ValueError('Nenhum funcionário com nome e e-mail preenchidos foi encontrado.')
        return people
    finally:
        workbook.close()


def prepare(excel, pdf):
    skipped_people = []
    people = read_people(excel, skipped_people)
    problems = []
    ignored_pages = []
    with Path(pdf).open('rb') as source:
        doc = PdfReader(source)
        for index, page in enumerate(doc.pages):
            text = page.extract_text() or ''
            # This receipt layout places the employee's value before its label
            # in PDF text order. Exact field comparison also distinguishes names
            # such as Ana Maria Silva and Ana Maria Silva Santos.
            fields = {normalize_name(n) for n in re.findall(r'([^\n]+)\nNome do Funcionário', text)}
            if len(fields) > 1:
                problems.append(f'Página {index + 1}: mais de um funcionário no recibo.')
                continue
            found = [name for name in fields if name in people] if fields else match_names(text, people)
            if not found:
                ignored_pages.append(index + 1)
            elif len(found) != 1:
                problems.append(f'Página {index + 1}: mais de um nome completo encontrado; correspondência ambígua.')
            else:
                people[found[0]]['pages'].append(index)
        for person in people.values():
            if not person['pages']:
                problems.append(f'{person["nome"]}: sem contracheque.')
        if problems:
            raise ValueError('Envio bloqueado. Corrija estas pendências:\n\n' + '\n'.join(problems))
        digest = hashlib.sha256(Path(pdf).read_bytes()).hexdigest()
        folder = BASE / 'lotes' / datetime.now().strftime('%Y%m%d_%H%M%S_%f')
        folder.mkdir(parents=True)
        for number, person in enumerate(people.values(), 1):
            target = folder / f'contracheque_{number:04d}.pdf'
            out = PdfWriter()
            for page in person['pages']:
                out.add_page(doc.pages[page])
            with target.open('wb') as destination:
                out.write(destination)
            person['file'] = str(target)
            person['status'] = 'Pendente'
        return people, folder, digest, {'ignored_pages': ignored_pages, 'without_email': skipped_people}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Departamento Pessoal • Contracheques')
        self.geometry('1040x720')
        self.people = None
        self.busy = False
        self.excel, self.pdf = tk.StringVar(), tk.StringVar()
        self.subject = tk.StringVar(value='Seu contracheque')
        self.account = tk.StringVar()
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='Envio de contracheques', font=('Segoe UI', 20, 'bold')).pack(anchor='w')
        ttk.Label(frame, text='Excel → conferir pelo nome completo → revisar os PDFs → criar rascunhos ou enviar pelo Outlook clássico').pack(anchor='w', pady=(5, 15))
        for label, var, types in [('Excel dos funcionários', self.excel, [('Excel', '*.xlsx')]), ('PDF dos contracheques', self.pdf, [('PDF', '*.pdf')])]:
            row = ttk.Frame(frame)
            row.pack(fill='x', pady=4)
            ttk.Label(row, text=label, width=26).pack(side='left')
            ttk.Entry(row, textvariable=var).pack(side='left', fill='x', expand=True)
            ttk.Button(row, text='Selecionar', command=lambda v=var, t=types: self.select(v, t)).pack(side='left', padx=8)
        ttk.Button(frame, text='1. Processar e conferir arquivos', command=self.process).pack(anchor='w', pady=10)
        self.table = ttk.Treeview(frame, columns=('nome', 'email', 'paginas', 'status'), show='headings', height=10)
        for key, title, width in [('nome', 'Funcionário', 240), ('email', 'Destinatário', 300), ('paginas', 'Páginas do original', 130), ('status', 'Situação', 180)]:
            self.table.heading(key, text=title)
            self.table.column(key, width=width)
        self.table.pack(fill='both', expand=True)
        self.table.bind('<Double-1>', self.open_pdf)
        ttk.Label(frame, text='Dê dois cliques em uma linha para abrir e revisar o PDF separado.').pack(anchor='w', pady=5)
        for title, var in [('Conta remetente (e-mail exato configurado no Outlook)', self.account), ('Assunto', self.subject)]:
            ttk.Label(frame, text=title).pack(anchor='w')
            ttk.Entry(frame, textvariable=var).pack(fill='x', pady=(0, 5))
        ttk.Label(frame, text='Mensagem (o nome do funcionário será acrescentado automaticamente)').pack(anchor='w')
        self.body = tk.Text(frame, height=3, font=('Segoe UI', 10))
        self.body.insert('1.0', 'Segue em anexo o seu contracheque.\nDepartamento Pessoal')
        self.body.pack(fill='x')
        actions = ttk.Frame(frame)
        actions.pack(fill='x', pady=10)
        ttk.Button(actions, text='2. Criar rascunhos no Outlook', command=lambda: self.deliver(False)).pack(side='left')
        ttk.Button(actions, text='3. Enviar e-mails agora', command=lambda: self.deliver(True)).pack(side='right')
        self.note = tk.StringVar(value='Aguardando os arquivos. Use o Excel com colunas: nome, email (CPF opcional).')
        ttk.Label(frame, textvariable=self.note).pack(anchor='w')
        self.protocol('WM_DELETE_WINDOW', self.close)

    def close(self):
        if self.busy:
            messagebox.showinfo('Operação em andamento', 'Aguarde a operação terminar antes de fechar.')
        else:
            self.destroy()

    def select(self, var, types):
        if self.busy:
            return
        path = filedialog.askopenfilename(filetypes=types)
        if path:
            var.set(path)
            self.people = None
            self.table.delete(*self.table.get_children())

    def process(self):
        if self.busy:
            return
        self.people = None
        self.table.delete(*self.table.get_children())
        try:
            self.people, self.folder, self.digest, self.summary = prepare(self.excel.get(), self.pdf.get())
            self.refresh()
            self.note.set(f'{len(self.people)} destinatário(s); {len(self.summary["ignored_pages"])} página(s) ignorada(s); {len(self.summary["without_email"])} cadastro(s) sem e-mail.')
            details = '\n'.join(self.summary['without_email'])
            ignored = ', '.join(map(str, self.summary['ignored_pages'])) or 'nenhuma'
            (self.folder / 'conferencia.txt').write_text(
                f'Destinatários: {len(self.people)}\nPáginas ignoradas: {ignored}\nCadastros sem e-mail: {len(self.summary["without_email"])}\n{details}\n', encoding='utf-8')
            messagebox.showinfo('Conferência do lote', self.note.get() + '\n\nSomente as pessoas exibidas na tabela receberão mensagens. Revise os PDFs antes de continuar.\nDetalhes em conferencia.txt na pasta do lote.')
        except Exception as exc:
            messagebox.showerror('Não foi possível processar', str(exc))

    def refresh(self):
        self.table.delete(*self.table.get_children())
        for name_key, p in self.people.items():
            self.table.insert('', 'end', iid=name_key, values=(p['nome'], p['email'], ', '.join(str(i + 1) for i in p['pages']), p['status']))

    def open_pdf(self, event=None):
        import os
        selected = self.table.selection()
        if selected and self.people:
            os.startfile(self.people[selected[0]]['file'])

    def deliver(self, send):
        if self.busy:
            return
        if not self.people:
            return messagebox.showinfo('Conferência necessária', 'Processe os arquivos primeiro.')
        account, subject, body = self.account.get().strip(), self.subject.get().strip(), self.body.get('1.0', 'end').strip()
        if not account or not subject:
            return messagebox.showerror('Campos obrigatórios', 'Preencha a conta remetente e o assunto.')
        if any(p['status'] != 'Pendente' for p in self.people.values()):
            return messagebox.showwarning('Lote já utilizado', 'Este lote já teve uma tentativa. Confira o relatório e o Outlook antes de preparar outro lote. Isso evita reenvios acidentais.')
        ledger = BASE / 'envios.jsonl'
        previous = []
        if ledger.exists():
            try:
                previous = [json.loads(line) for line in ledger.read_text(encoding='utf-8').splitlines() if line.strip()]
                if any(not isinstance(record, dict) for record in previous):
                    raise ValueError('Formato inválido no histórico.')
            except Exception as exc:
                return messagebox.showerror('Histórico não pôde ser lido', f'Confira envios.jsonl antes de continuar: {exc}')
        eligible, blocked = split_previous_attempts(self.people, self.digest, previous)
        if not eligible:
            for key in blocked:
                self.people[key]['status'] = 'Tentativa anterior; conferir Outlook'
            self.refresh()
            return messagebox.showwarning('Destinatários já processados', 'Todos os destinatários deste lote já tiveram tentativa de envio deste PDF. Confira Itens Enviados, Caixa de Saída e o relatório. O histórico não confirma entrega; nenhum reenvio foi feito.')
        verb = 'ENVIAR AGORA' if send else 'CRIAR RASCUNHOS'
        warning = f'\n\n{len(blocked)} destinatário(s) com tentativa anterior serão ignorados. Confira-os no Outlook.' if blocked else ''
        if not messagebox.askyesno('Confirmar operação', f'{verb} para {len(eligible)} funcionário(s) pela conta {account}?{warning}\n\nConfirme que revisou os destinatários e os PDFs. Anexos contêm dados pessoais.\n\nRascunhos existentes não serão reutilizados pelo envio direto.'):
            return
        self.active_keys = set(eligible)
        for key in blocked:
            self.people[key]['status'] = 'Tentativa anterior; conferir Outlook'
        self.refresh()
        self.busy = True
        self.note.set('Operação em andamento. Aguarde…')
        threading.Thread(target=self.worker, args=(send, account, subject, body), daemon=True).start()

    def worker(self, send, address, subject, body):
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        error = None
        try:
            outlook = win32com.client.Dispatch('Outlook.Application')
            accounts = [a for a in outlook.Session.Accounts if a.SmtpAddress.lower() == address.lower()]
            if len(accounts) != 1:
                raise ValueError('Conta não encontrada no Outlook clássico. Abra o Outlook e confira o endereço remetente.')
            for name_key, person in self.people.items():
                if name_key not in self.active_keys:
                    continue
                mail = outlook.CreateItem(0)
                mail.SendUsingAccount = accounts[0]
                mail.To = person['email']
                mail.Subject = subject
                mail.Body = f'Olá, {person["nome"]}.\n\n{body}'
                mail.Attachments.Add(person['file'])
                if not mail.Recipients.ResolveAll():
                    raise ValueError(f'Destinatário não resolvido: {person["email"]}')
                # Record before Send: even a crash/ambiguous Outlook failure blocks automatic retries.
                if send:
                    with (BASE / 'envios.jsonl').open('a', encoding='utf-8') as log:
                        log.write(json.dumps({'data': datetime.now().isoformat(), 'pdf': self.digest, 'email': person['email'], 'estado': 'tentativa'}, ensure_ascii=False) + '\n')
                    person['status'] = 'Envio iniciado'
                    mail.Send()
                    person['status'] = 'Entregue ao Outlook'
                else:
                    person['status'] = 'Criação iniciada'
                    mail.Save()
                    person['status'] = 'Rascunho criado'
        except Exception as exc:
            error = str(exc)
        finally:
            pythoncom.CoUninitialize()
            try:
                with (self.folder / 'relatorio.csv').open('w', newline='', encoding='utf-8-sig') as file:
                    writer = csv.writer(file, delimiter=';')
                    writer.writerow(['Nome', 'Email', 'Paginas', 'Situacao'])
                    for p in self.people.values():
                        # Avoid formula interpretation when the report is opened in Excel.
                        safe = lambda s: "'" + s if str(s).startswith(('=', '+', '-', '@')) else s
                        writer.writerow([safe(p['nome']), safe(p['email']), ','.join(str(i + 1) for i in p['pages']), p['status']])
            except Exception as exc:
                error = f'{error or ""}\nFalha ao gravar relatório: {exc}'
            self.after(0, lambda e=error: self.finished(e))

    def finished(self, error):
        self.busy = False
        self.refresh()
        self.note.set('Operação interrompida. Confira o relatório.' if error else 'Concluído. Relatório salvo na pasta do lote.')
        if error:
            messagebox.showerror('Operação interrompida', error + '\n\nConfira os itens já processados no Outlook antes de tentar novamente.')
        else:
            messagebox.showinfo('Concluído', 'Operação concluída. Envios foram entregues ao Outlook; a entrega final depende do servidor da empresa.\n\nRelatório: ' + str(self.folder / 'relatorio.csv'))


if __name__ == '__main__':
    App().mainloop()
