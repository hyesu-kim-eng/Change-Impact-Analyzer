"""Change Impact Analyzer — local desktop prototype (Excel / text PDF)."""
import csv
import re
import sys
import traceback
from collections import defaultdict
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import fitz

TAG_RE = re.compile(r'(?<![A-Za-z0-9])(?:[A-Z0-9]{2,8}-){1,4}[A-Z0-9]{2,12}(?![A-Za-z0-9])', re.I)
RULES = {
    'signal': ('I/O List', 'Loop Diagram', 'Cable Schedule', 'DCS Configuration', 'FAT Procedure'),
    'output': ('I/O List', 'Loop Diagram', 'Cable Schedule', 'DCS Configuration'),
    'range': ('Data Sheet', 'DCS Scaling', 'Alarm / Interlock', 'FAT Procedure'),
    'setpoint': ('Cause & Effect', 'Logic Diagram', 'DCS Configuration', 'FAT Procedure'),
    'fail': ('Cause & Effect', 'Logic Diagram', 'Actuator Data Sheet', 'FAT Procedure'),
    'location': ('Instrument Layout', 'Cable Routing', 'JB Schedule'),
    'type': ('Instrument Index', 'Data Sheet', 'I/O List', 'Component Mapping'),
    'power': ('Power Distribution', 'I/O List', 'Wiring Diagram'),
    'cable': ('Cable Schedule', 'Loop Diagram', 'JB Wiring'),
    'alarm': ('Alarm List', 'Cause & Effect', 'DCS Configuration'),
}
TAG_COLUMNS = ('tag no', 'tag number', 'tagno', 'tag_number', 'instrument tag', 'tag', 'loop tag')

def norm(x):
    if pd.isna(x): return ''
    return re.sub(r'\s+', ' ', str(x).strip())

def key(x):
    return re.sub(r'[^a-z0-9]', '', str(x).lower())

def detect_tag_column(cols):
    for wanted in TAG_COLUMNS:
        for c in cols:
            if key(c) == key(wanted): return c
    for c in cols:
        if 'tag' in key(c) and not any(w in key(c) for w in ('remark', 'status')): return c
    return None

def load_excel(path):
    book = pd.read_excel(path, sheet_name=None, dtype=str, header=None)
    records = defaultdict(dict)
    notes = []
    for sheet, raw in book.items():
        if raw.empty: continue
        detected = None
        for i in range(min(25, len(raw))):
            cols = [norm(x) for x in raw.iloc[i].tolist()]
            tag_col = detect_tag_column(cols)
            if tag_col is not None:
                detected = (i, cols.index(tag_col), cols)
                break
        if not detected:
            notes.append(f'{sheet}: no recognizable Tag header in first 25 rows')
            continue
        head, tag_idx, cols = detected
        for ri in range(head + 1, len(raw)):
            vals = raw.iloc[ri].tolist()
            tag = norm(vals[tag_idx]).upper()
            if not tag or tag in ('NAN', 'NONE'): continue
            if len(tag) > 90: continue
            for ci, col in enumerate(cols):
                if ci == tag_idx or not col or col.lower() in ('nan', 'none'): continue
                value = norm(vals[ci])
                if value:
                    records[tag][f'{sheet} / {col}'] = value
    return dict(records), notes

def load_pdf(path):
    doc = fitz.open(path)
    records = defaultdict(dict)
    notes = []
    for pi, page in enumerate(doc):
        text = page.get_text('text')
        if not text.strip():
            notes.append(f'Page {pi+1}: no selectable text (scanned PDF may need OCR)')
            continue
        tags = sorted(set(t.upper() for t in TAG_RE.findall(text)))
        for tag in tags:
            lines = text.splitlines()
            matches = [i for i, line in enumerate(lines) if tag in line.upper()]
            snippet = ' | '.join(' '.join(lines[max(0, i-1):min(len(lines),i+2)])[:250] for i in matches[:3])
            records[tag][f'PDF / Page {pi+1}'] = snippet or 'Tag mentioned'
    notes.append('PDF comparison uses tag occurrence/context, not verified datasheet field extraction.')
    return dict(records), notes

def load(path):
    ext = Path(path).suffix.lower()
    if ext in ('.xlsx', '.xlsm', '.xls'): return load_excel(path)
    if ext == '.pdf': return load_pdf(path)
    raise ValueError(f'Unsupported file: {ext}')

def classify(field):
    k = key(field)
    found = []
    for needle, targets in RULES.items():
        if needle in k: found.extend(targets)
    return sorted(set(found)) or ['Engineering Review (rule not mapped)']

def analyze(old, new):
    results = []
    for tag in sorted(set(old) | set(new)):
        a, b = old.get(tag), new.get(tag)
        if a is None or b is None:
            status = 'ADDED' if a is None else 'REMOVED'
            results.append((tag, status, '(tag)', '', '', 'All related documents', 'Verify scope and linked records'))
            continue
        for field in sorted(set(a) | set(b)):
            va, vb = a.get(field, ''), b.get(field, '')
            if va != vb:
                for target in classify(field):
                    results.append((tag, 'CHANGED', field, va, vb, target, 'REVIEW - rule-based potential impact'))
    return results

HEADERS = ['Tag', 'Change Status', 'Source Field', 'Old Value', 'New Value', 'Potentially Affected Item', 'Assessment']

class App(tk.Tk):
    BG='#F8F6FF'; SIDE='#F0EAFB'; PANEL='#FFFFFF'; PANEL2='#F2ECFC'; BORDER='#E5DCF3'
    TEXT='#49336D'; MUTED='#8C7BA5'; ACCENT='#8D6AC8'; GREEN='#4A9E85'; AMBER='#C18A3D'; RED='#D5798B'
    def __init__(self):
        super().__init__()
        self.title('Engineering Change Impact Analyzer | Lavender Studio')
        self.geometry('1320x820'); self.minsize(1050,650); self.configure(bg=self.BG)
        self.old=tk.StringVar(); self.new=tk.StringVar()
        self.status=tk.StringVar(value='Ready  •  Select previous and new revision files to begin')
        self.search=tk.StringVar(); self.filter=tk.StringVar(value='ALL')
        self.results=[]; self.notes=[]; self.stats={}
        self._styles(); self._ui()
        self.search.trace_add('write', lambda *_: self.refresh())

    def _styles(self):
        st=ttk.Style(self); st.theme_use('clam')
        st.configure('Lavender.Treeview', background=self.PANEL, fieldbackground=self.PANEL,
                     foreground=self.TEXT, rowheight=35, borderwidth=0, font=('Segoe UI',10))
        st.configure('Lavender.Treeview.Heading', background=self.PANEL2, foreground='#655080',
                     font=('Segoe UI',9,'bold'), relief='flat', padding=(9,13))
        st.map('Lavender.Treeview', background=[('selected','#E9DFFF')], foreground=[('selected','#49336D')])
        st.map('Lavender.Treeview.Heading', background=[('active','#E9DFFF')])
        st.configure('Lavender.Vertical.TScrollbar', background=self.PANEL2, troughcolor=self.BG, borderwidth=0)
        st.configure('Lavender.Horizontal.TScrollbar', background=self.PANEL2, troughcolor=self.BG, borderwidth=0)

    def label(self,parent,text,size=10,color=None,bold=False,**kw):
        return tk.Label(parent,text=text,font=('Segoe UI',size,'bold' if bold else 'normal'),
                        fg=color or self.TEXT,bg=parent.cget('bg'),**kw)

    def button(self,parent,text,cmd,primary=False,width=None):
        bg=self.ACCENT if primary else self.PANEL2
        fg='#FFFFFF' if primary else self.TEXT
        b=tk.Button(parent,text=text,command=cmd,bg=bg,fg=fg,activebackground='#7953B5' if primary else '#E9DFFF',
                    activeforeground=fg,relief='flat',bd=0,padx=18,pady=11,cursor='hand2',
                    font=('Segoe UI',10,'bold'),width=width)
        return b

    def _ui(self):
        sidebar=tk.Frame(self,bg=self.SIDE,width=230); sidebar.pack(side='left',fill='y'); sidebar.pack_propagate(False)
        tk.Frame(sidebar,bg=self.ACCENT,height=5).pack(fill='x')
        self.label(sidebar,'✦  ENGINEERING',13,self.TEXT,True).pack(anchor='w',padx=23,pady=(29,3))
        self.label(sidebar,'AUTOMATION SUITE',9,self.MUTED,True).pack(anchor='w',padx=24)
        tk.Frame(sidebar,bg=self.BORDER,height=1).pack(fill='x',padx=20,pady=(27,23))
        nav=tk.Frame(sidebar,bg='#E5D9F8'); nav.pack(fill='x',padx=12)
        self.label(nav,'  ◉   Change Impact',11,self.TEXT,True).pack(anchor='w',padx=9,pady=13)
        self.label(sidebar,'WORKSPACE',9,self.MUTED,True).pack(anchor='w',padx=24,pady=(28,10))
        self.label(sidebar,'  01   Revision comparison',10,'#6D5A87').pack(anchor='w',padx=20,pady=9)
        self.label(sidebar,'  02   Impact assessment',10,'#6D5A87').pack(anchor='w',padx=20,pady=9)
        self.label(sidebar,'  03   Export report',10,'#6D5A87').pack(anchor='w',padx=20,pady=9)
        tk.Frame(sidebar,bg=self.SIDE).pack(expand=True)
        self.label(sidebar,'●  LOCAL PROCESSING',9,self.GREEN,True).pack(anchor='w',padx=24,pady=(0,8))
        self.label(sidebar,'v1.2  •  Excel / Text PDF',9,self.MUTED).pack(anchor='w',padx=24,pady=(0,27))

        main=tk.Frame(self,bg=self.BG); main.pack(side='left',fill='both',expand=True)
        head=tk.Frame(main,bg=self.BG); head.pack(fill='x',padx=30,pady=(25,17))
        self.label(head,'Change Impact Analyzer',23,self.TEXT,True).pack(anchor='w')
        self.label(head,'Track engineering revisions, identify potential impacts, and prioritize review.',10,self.MUTED).pack(anchor='w',pady=(5,0))

        files=tk.Frame(main,bg=self.PANEL,highlightbackground=self.BORDER,highlightthickness=1)
        files.pack(fill='x',padx=30,pady=(0,17))
        top=tk.Frame(files,bg=self.PANEL); top.pack(fill='x',padx=19,pady=(17,10))
        self.label(top,'REVISION COMPARISON',10,self.TEXT,True).pack(side='left')
        self.label(top,'Supported: .xlsx  .xlsm  .xls  .pdf',9,self.MUTED).pack(side='right')
        for name,var in [('PREVIOUS REVISION',self.old),('NEW REVISION',self.new)]:
            line=tk.Frame(files,bg=self.PANEL); line.pack(fill='x',padx=19,pady=5)
            self.label(line,name,9,self.MUTED,True,width=19,anchor='w').pack(side='left')
            e=tk.Entry(line,textvariable=var,bg=self.BG,fg=self.TEXT,insertbackground=self.TEXT,
                       relief='flat',font=('Segoe UI',10),highlightthickness=1,highlightbackground=self.BORDER,
                       highlightcolor=self.ACCENT)
            e.pack(side='left',fill='x',expand=True,ipady=10,padx=(3,10))
            self.button(line,'Browse',lambda v=var:self.browse(v)).pack(side='right')
        actions=tk.Frame(files,bg=self.PANEL); actions.pack(fill='x',padx=19,pady=(12,18))
        self.button(actions,'▶  Analyze Changes',self.run,True).pack(side='left')
        self.button(actions,'↓  Export CSV',self.export).pack(side='left',padx=10)
        self.label(actions,'Rule-based impact assessment  •  Local processing',9,self.MUTED).pack(side='right')

        cards=tk.Frame(main,bg=self.BG); cards.pack(fill='x',padx=30,pady=(0,17))
        self.card_labels={}
        for idx,(name,color) in enumerate([('TOTAL CHANGES',self.ACCENT),('CHANGED TAGS',self.AMBER),('ADDED TAGS',self.GREEN),('REMOVED TAGS',self.RED)]):
            c=tk.Frame(cards,bg=self.PANEL,highlightbackground=self.BORDER,highlightthickness=1)
            c.pack(side='left',fill='x',expand=True,padx=(0,10 if idx<3 else 0))
            self.label(c,name,9,self.MUTED,True).pack(anchor='w',padx=17,pady=(14,6))
            num=self.label(c,'—',23,color,True); num.pack(anchor='w',padx=17,pady=(0,12))
            self.card_labels[name]=num

        result=tk.Frame(main,bg=self.PANEL,highlightbackground=self.BORDER,highlightthickness=1)
        result.pack(fill='both',expand=True,padx=30,pady=(0,12))
        bar=tk.Frame(result,bg=self.PANEL); bar.pack(fill='x',padx=16,pady=(14,12))
        self.label(bar,'CHANGE ANALYSIS RESULTS',11,self.TEXT,True).pack(side='left')
        tk.Label(bar,text='   ',bg=self.PANEL).pack(side='left')
        self.count_label=self.label(bar,'0 rows',9,self.MUTED); self.count_label.pack(side='left')
        self.button(bar,'View Details',self.details).pack(side='right',padx=(9,0))
        search=tk.Entry(bar,textvariable=self.search,bg=self.BG,fg=self.TEXT,insertbackground=self.TEXT,
                        relief='flat',width=23,font=('Segoe UI',10),highlightthickness=1,
                        highlightbackground=self.BORDER,highlightcolor=self.ACCENT)
        search.pack(side='right',ipady=9,padx=(9,0))
        self.label(bar,'Search:',9,self.MUTED).pack(side='right')
        for name in ['ALL','CHANGED','ADDED','REMOVED']:
            b=tk.Radiobutton(bar,text=name,variable=self.filter,value=name,command=self.refresh,
                bg=self.PANEL,fg=self.MUTED,selectcolor='#E9DFFF',activebackground=self.PANEL,
                activeforeground=self.ACCENT,indicatoron=False,relief='flat',bd=0,padx=8,pady=7,
                font=('Segoe UI',8,'bold'),cursor='hand2')
            b.pack(side='right',padx=2)
        table=tk.Frame(result,bg=self.PANEL); table.pack(fill='both',expand=True,padx=13,pady=(0,13))
        self.tree=ttk.Treeview(table,columns=HEADERS,show='headings',style='Lavender.Treeview',selectmode='browse')
        for name,width in zip(HEADERS,[135,110,190,175,175,220,245]):
            self.tree.heading(name,text=name); self.tree.column(name,width=width,minwidth=90)
        self.tree.tag_configure('ADDED',foreground=self.GREEN)
        self.tree.tag_configure('REMOVED',foreground=self.RED)
        self.tree.tag_configure('CHANGED',foreground='#7953B5')
        sy=ttk.Scrollbar(table,orient='vertical',command=self.tree.yview,style='Lavender.Vertical.TScrollbar')
        sx=ttk.Scrollbar(table,orient='horizontal',command=self.tree.xview,style='Lavender.Horizontal.TScrollbar')
        self.tree.configure(yscrollcommand=sy.set,xscrollcommand=sx.set)
        self.tree.grid(row=0,column=0,sticky='nsew'); sy.grid(row=0,column=1,sticky='ns'); sx.grid(row=1,column=0,sticky='ew')
        table.rowconfigure(0,weight=1); table.columnconfigure(0,weight=1)
        self.tree.bind('<Double-1>',lambda _:self.details())
        foot=tk.Frame(main,bg=self.BG); foot.pack(fill='x',padx=31,pady=(0,15))
        self.label(foot,'ⓘ  Results indicate potential impacts, not confirmed design changes. Scanned PDFs are not supported.',9,self.MUTED).pack(anchor='w')
        self.label(foot,'STATUS',8,self.ACCENT,True).pack(side='left',pady=(9,0))
        self.label(foot,'  ',9).pack(side='left')
        self.label(foot,'',9).pack_forget() if False else None
        tk.Label(foot,textvariable=self.status,bg=self.BG,fg=self.MUTED,font=('Segoe UI',9)).pack(side='left',pady=(9,0))

    def browse(self,variable):
        p=filedialog.askopenfilename(filetypes=[('Excel and PDF','*.xlsx *.xlsm *.xls *.pdf'),('All files','*.*')])
        if p:variable.set(p)

    def refresh(self):
        if not hasattr(self,'tree'):return
        self.tree.delete(*self.tree.get_children())
        q=self.search.get().strip().lower(); f=self.filter.get()
        shown=0
        for row in self.results:
            if f!='ALL' and row[1]!=f:continue
            if q and q not in ' '.join(map(str,row)).lower():continue
            if shown>=20000:break
            self.tree.insert('','end',values=row,tags=(row[1],)); shown+=1
        self.count_label.config(text=f'{shown:,} displayed / {len(self.results):,} total')

    def run(self):
        if not Path(self.old.get()).is_file() or not Path(self.new.get()).is_file():
            messagebox.showwarning('Missing files','Select valid previous and new revision files.');return
        self.status.set('Analyzing revisions…');self.update_idletasks()
        try:
            old,n1=load(self.old.get());new,n2=load(self.new.get())
            self.notes=n1+n2;self.results=analyze(old,new)
            changes={r[0] for r in self.results if r[1]=='CHANGED'}
            added={r[0] for r in self.results if r[1]=='ADDED'}
            removed={r[0] for r in self.results if r[1]=='REMOVED'}
            for k,v in [('TOTAL CHANGES',len(self.results)),('CHANGED TAGS',len(changes)),('ADDED TAGS',len(added)),('REMOVED TAGS',len(removed))]:
                self.card_labels[k].config(text=f'{v:,}')
            self.refresh()
            self.status.set(f'Complete  •  {len(old):,} previous tags  /  {len(new):,} new tags')
            if self.notes:messagebox.showinfo('Extraction notes','\n'.join(self.notes[:15]))
        except Exception as e:
            self.status.set('Analysis failed');messagebox.showerror('Analysis failed',f'{e}\n\n{traceback.format_exc()[-1200:]}')

    def details(self):
        selected=self.tree.selection()
        if not selected:
            messagebox.showinfo('Select a result','Select a result row first.');return
        values=self.tree.item(selected[0],'values')
        win=tk.Toplevel(self);win.title('Impact Review Detail');win.geometry('700x450');win.configure(bg=self.BG)
        self.label(win,'Impact Review Detail',19,self.TEXT,True).pack(anchor='w',padx=24,pady=(22,15))
        for h,v in zip(HEADERS,values):
            r=tk.Frame(win,bg=self.PANEL);r.pack(fill='x',padx=24,pady=3)
            self.label(r,h,9,self.MUTED,True,width=24,anchor='w').pack(side='left',padx=12,pady=9)
            self.label(r,str(v),10,self.TEXT,wraplength=410,justify='left',anchor='w').pack(side='left',padx=5)
        self.button(win,'Close',win.destroy,True).pack(anchor='e',padx=24,pady=15)

    def export(self):
        if not self.results:
            messagebox.showinfo('Nothing to export','Run analysis first (or no changes were found).');return
        p=filedialog.asksaveasfilename(defaultextension='.csv',filetypes=[('CSV report','*.csv')],initialfile='change_impact_report.csv')
        if not p:return
        with open(p,'w',encoding='utf-8-sig',newline='') as f:
            w=csv.writer(f);w.writerow(HEADERS);w.writerows(self.results)
        messagebox.showinfo('Report exported',p)

if __name__=='__main__':
    App().mainloop()
