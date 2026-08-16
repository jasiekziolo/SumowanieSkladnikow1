import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import threading

try:
    from openpyxl import load_workbook, Workbook
except ImportError:
    load_workbook = Workbook = None

APP_TITLE = "Sumowanie składników Excel"

class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry("1000x650")
        self.root.minsize(800, 500)
        self.rows = []
        self.source = None
        self.setup_style()
        self.build()

    def setup_style(self):
        s = ttk.Style()
        try: s.theme_use("vista")
        except tk.TclError: pass
        s.configure("Title.TLabel", font=("Segoe UI", 18, "bold"))
        s.configure("Action.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))
        s.configure("Treeview", font=("Segoe UI", 10), rowheight=28)
        s.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"))

    def build(self):
        outer = ttk.Frame(self.root, padding=18); outer.pack(fill="both", expand=True)
        top = ttk.Frame(outer); top.pack(fill="x", pady=(0,12))
        ttk.Label(top, text=APP_TITLE, style="Title.TLabel").pack(side="left")
        ttk.Button(top, text="Wybierz plik Excel", command=self.choose, style="Action.TButton").pack(side="right")

        info = tk.Frame(outer, bd=1, relief="solid", padx=18, pady=16)
        info.pack(fill="x", pady=(0,12))
        tk.Label(info, text="Wczytaj plik .xlsx", font=("Segoe UI",12,"bold")).pack()
        tk.Label(info, text="Kolumny: B – Numer składnika   D – Opis   E – Ilość   F – Jednostka\n"
                 "Pozycje są grupowane po numerze, opisie i jednostce, a ilości są sumowane.",
                 font=("Segoe UI",10), justify="center").pack(pady=(5,0))

        self.status = tk.StringVar(value="Gotowe. Wybierz plik Excel.")
        ttk.Label(outer, textvariable=self.status).pack(fill="x", pady=(0,8))

        frame = ttk.Frame(outer); frame.pack(fill="both", expand=True)
        cols=("nr","opis","ilosc","jednostka")
        self.tree=ttk.Treeview(frame, columns=cols, show="headings")
        headings={"nr":"Numer składnika","opis":"Opis składnika","ilosc":"Ilość","jednostka":"Jednostka"}
        widths={"nr":160,"opis":500,"ilosc":130,"jednostka":120}
        for c in cols:
            self.tree.heading(c,text=headings[c]); self.tree.column(c,width=widths[c],anchor="w" if c=="opis" else "center")
        self.tree.column("ilosc",anchor="e")
        ys=ttk.Scrollbar(frame,orient="vertical",command=self.tree.yview)
        xs=ttk.Scrollbar(frame,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set,xscrollcommand=xs.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); ys.grid(row=0,column=1,sticky="ns"); xs.grid(row=1,column=0,sticky="ew")
        frame.rowconfigure(0,weight=1); frame.columnconfigure(0,weight=1)

        bottom=ttk.Frame(outer); bottom.pack(fill="x",pady=(12,0))
        self.count=tk.StringVar(value="Liczba pozycji: 0")
        ttk.Label(bottom,textvariable=self.count).pack(side="left")
        self.save=ttk.Button(bottom,text="Zapisz jako Excel",command=self.save_file,style="Action.TButton",state="disabled")
        self.save.pack(side="right")

    def choose(self):
        if load_workbook is None:
            messagebox.showerror("Błąd","Brak biblioteki openpyxl w tej wersji programu.")
            return
        p=filedialog.askopenfilename(title="Wybierz plik Excel",filetypes=[("Pliki Excel","*.xlsx")])
        if p: self.process(p)

    def process(self,path):
        self.status.set("Przetwarzanie pliku..."); self.save.config(state="disabled")
        threading.Thread(target=self.worker,args=(path,),daemon=True).start()

    def worker(self,path):
        try:
            wb=load_workbook(path,read_only=True,data_only=True)
            ws=wb.active
            h={c:ws[f"{c}1"].value for c in "BDEF"}
            aliases={"B":{"numer składnika","numer skladnika","nr składnika","nr skladnika"},
                     "D":{"opis składnika","opis skladnika"},"E":{"ilość","ilosc"},"F":{"jednostka","jednostka miary"}}
            header=all(h[c] is not None and str(h[c]).strip().lower() in aliases[c] for c in aliases)
            start=2 if header else 1
            grouped=defaultdict(Decimal); errors=[]
            for n,row in enumerate(ws.iter_rows(min_row=start,min_col=2,max_col=6,values_only=True),start):
                nr,_,opis,ilosc,jed=row
                if all(v is None or str(v).strip()=="" for v in (nr,opis,ilosc,jed)): continue
                if nr is None or str(nr).strip()=="": errors.append(f"Wiersz {n}: brak Numeru składnika."); continue
                if ilosc is None or str(ilosc).strip()=="": errors.append(f"Wiersz {n}: brak Ilości."); continue
                try:
                    if isinstance(ilosc,bool): raise InvalidOperation
                    q=Decimal(str(ilosc).replace(",",".").strip())
                except (InvalidOperation,ValueError):
                    errors.append(f"Wiersz {n}: '{ilosc}' w kolumnie E nie jest liczbą."); continue
                key=(str(nr).strip(),str(opis or "").strip(),str(jed or "").strip())
                grouped[key]+=q
            wb.close()
            result=[(k[0],k[1],v,k[2]) for k,v in grouped.items()]
            result.sort(key=lambda x:x[0])
            self.root.after(0,lambda:self.show_result(path,result,errors))
        except Exception as e:
            self.root.after(0,lambda:messagebox.showerror("Błąd odczytu",str(e)))
            self.root.after(0,lambda:self.status.set("Wystąpił błąd."))

    @staticmethod
    def fmt(x):
        if x==x.to_integral(): return str(int(x))
        return format(x.normalize(),"f").rstrip("0").rstrip(".")

    def show_result(self,path,result,errors):
        self.source=path; self.rows=result
        for i in self.tree.get_children(): self.tree.delete(i)
        for r in result: self.tree.insert("","end",values=(r[0],r[1],self.fmt(r[2]),r[3]))
        self.count.set(f"Liczba pozycji po zsumowaniu: {len(result)}")
        self.save.config(state="normal" if result else "disabled")
        self.status.set(f"Gotowe: {len(result)} pozycji.")
        if errors:
            msg="\n".join(errors[:15])
            if len(errors)>15: msg += f"\n... oraz {len(errors)-15} kolejnych."
            messagebox.showwarning("Błędne dane","Część wierszy pominięto:\n\n"+msg)

    def save_file(self):
        if not self.rows: return
        name=(Path(self.source).stem+"_zsumowane.xlsx") if self.source else "zsumowane_skladniki.xlsx"
        p=filedialog.asksaveasfilename(title="Zapisz wynik",defaultextension=".xlsx",initialfile=name,filetypes=[("Plik Excel","*.xlsx")])
        if not p:return
        try:
            wb=Workbook(); ws=wb.active; ws.title="Zsumowane składniki"
            ws.append(["Numer składnika","Opis składnika","Ilość","Jednostka"])
            for nr,opis,q,jed in self.rows: ws.append([nr,opis,float(q),jed])
            ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
            for c,w in {"A":20,"B":55,"C":15,"D":15}.items(): ws.column_dimensions[c].width=w
            wb.save(p); self.status.set("Zapisano plik."); messagebox.showinfo("Zapisano",f"Plik zapisany:\n\n{p}")
        except PermissionError: messagebox.showerror("Brak dostępu","Zamknij plik w Excelu i spróbuj ponownie.")
        except Exception as e: messagebox.showerror("Błąd zapisu",str(e))

def main():
    root=tk.Tk(); App(root); root.mainloop()
if __name__=="__main__": main()
