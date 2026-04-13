import os
import threading
import time
import copy
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.cell.cell import MergedCell
from docx import Document
from pptx import Presentation
from deep_translator import GoogleTranslator
import customtkinter as ctk
from tkinter import filedialog, messagebox
import random
import gc
import subprocess

try:
    from win32com.client import Dispatch
    import win32com.client as win32
except ImportError:
    win32 = None

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

romantic_notes = [
    "Traduciendo con amor...",
    "Haciendo tu vida un poquito más fácil...",
    "Casi listo, como nosotros...",
    "Unos segundos más para el mejor traductor...",
    "Preparando algo especial para ti...",
    "Relájate, yo me encargo de los idiomas hoy",
    "Un documento menos, un beso más",
    "Viste lo rápido que lo hago? Solo para impresionarte...",
    "Traducir es fácil, lo difícil es no distraerme pensando en ti",
    "Si este progreso fuera una cita, ya estaríamos en el postre",
    "Mírame procesar... Sé que te gusta",
    "Ocupada? No importa, yo sigo aquí dándolo todo por ti",
    "Traduciendo... Porque Google no te quiere como yo",
    "Analizando archivos... Ojalá entenderte a ti fuera así de simple",
    "Haciendo magia. No me preguntes cómo, solo disfrútalo",
    "Error 404: Corazón no encontrado... No mentiras, aquí estoy",
    "Me encanta cuando me dejas trabajar así de duro...",
    "Si vieras cómo se calienta mi CPU cuando me usas...",
    "Traduciendo... Quiero darte un archivo grande",
    "Quieres ver mi código fuente o prefieres seguir esperando?",
    "Tocando este documento... Y no es lo único que quiero tocar",
    "Estoy guardando cada movimiento, aprender para la próxima vez que me uses 😉",
]

final_love_notes = [
    # --- Románticos & Dulces ---
    "✨ ¡Todo listo! ✨\n\nEspero que esto te ahorre\nmucho tiempo hoy.\n\nTe quiero muchoooo!.",
    "Hecho con amor para la persona\nque traduce mi mundo.\n\n❤︎",
    "Misión cumplida.\n\nAhora ve a descansar,\nte lo has ganado.",
    "Cada palabra fue traducida\npensando en ti.\n\nDisfruta tu tiempo libre.",
    "Traducción terminada.\n\nMi código solo tiene sentido\nsi es para ayudarte.",

    # --- Juguetones & Coquetos ---
    "¡Listo!\n\nMe debes un beso por cada\npágina traducida. Haz la cuenta.",
    "Terminé mi trabajo...\n\nAhora qué vamos a hacer\ncon todo ese tiempo libre? 😉",
    "Documentos listos.\n\nYa te dije que te ves\nincreíble hoy?",
    "Traducción exitosa.\n\nSoy el mejor asistente que\npodrías tener, admítelo.",

    # --- Sarcásticos & Divertidos ---
    "Listo.\n\nSi hay algún error, fue culpa de Google.\nSi está perfecto, fue gracias a mi.",
    "Ya terminé.\n\nPor favor, no me pidas que traduzca\ntus sentimientos, eso es mas dificil.",
    "Archivo procesado.\n\nHe trabajado más que tú en la última hora,\npero no dire nada. 😉",

    # --- Picantes / Atrevidos (Dirty-ish) ---
    "Terminé de procesar!\n\nMe encanta cuando me usas\nhasta que termino todo",
    "Traducción completada.\n\nAhora que cerraste el archivo...\nPor qué no descansamos juntos? 🔥",
    "Todo listo.\n\nSi crees que soy rápido trabajando,\nespera a que apagues la computadora.",
    "Hecho.\n\nMe dejaste el CPU a 90 grados...\nVas a hacer algo para enfriarme?",
    "Quiero darte\ntareas tan largas y pesadas... 😈"
]

class DocTranslatorPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("With Love ❤︎ - Doc Translator")

        self.geometry("1000x820")
        
        self.is_cancelled = False
        self.translator = GoogleTranslator(source='auto', target='es')
        self.files_path_set = set()
        self.files_to_process = []
        
        self.total_units = 0
        self.done_units = 0
        self.start_time = 0

        self._build_ui()
        self._animate_heart()

    def _animate_heart(self):
        hearts = ["❤︎", "♥", "♡"]
        idx = 0
        def run():
            while True:
                try:
                    h = hearts[idx % len(hearts)]
                    self.title(f"With Love {h} - Doc Translator")
                    idx += 1
                    time.sleep(0.8)
                except: break
        threading.Thread(target=run, daemon=True).start()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self.head = ctk.CTkLabel(self, text="TRANSLATOR FOR THE BEST TRANSLATOR", font=ctk.CTkFont(size=32, weight="bold"))
        self.head.grid(row=0, column=0, pady=(40, 20))

        self.up_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.up_frame.grid(row=1, column=0, padx=20, pady=10)

        self.btn_fold = ctk.CTkButton(self.up_frame, text="📂 Carpeta", command=self.add_folder, fg_color="#27ae60", width=150)
        self.btn_fold.grid(row=0, column=0, padx=10)

        self.btn_one = ctk.CTkButton(self.up_frame, text="📄 Archivo", command=self.add_file, fg_color="#2980b9", width=150)
        self.btn_one.grid(row=0, column=1, padx=10)

        self.btn_clr = ctk.CTkButton(self.up_frame, text="🗑️ Limpiar", command=self.clear, fg_color="#7f8c8d", width=150)
        self.btn_clr.grid(row=0, column=2, padx=10)

        self.list_frame = ctk.CTkScrollableFrame(self, label_text="Cola de Documentos (Word, Excel, PPT)", label_font=ctk.CTkFont(weight="bold"))
        self.list_frame.grid(row=2, column=0, padx=50, pady=20, sticky="nsew")

        self.bot_panel = ctk.CTkFrame(self, corner_radius=20, fg_color="#1a1a1a")
        self.bot_panel.grid(row=3, column=0, padx=40, pady=30, sticky="ew")
        self.bot_panel.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(self.bot_panel, text="Listo para procesar", font=ctk.CTkFont(size=16))
        self.lbl_status.grid(row=0, column=0, pady=(20, 5))

        self.lbl_msg = ctk.CTkLabel(self.bot_panel, text="", font=ctk.CTkFont(size=16))
        self.lbl_msg.grid(row=1, column=0, pady=(0, 5))

        self.lbl_eta = ctk.CTkLabel(self.bot_panel, text="ETA: --:--", text_color="#95a5a6")
        self.lbl_eta.grid(row=2, column=0, pady=(0, 10))

        self.bar = ctk.CTkProgressBar(self.bot_panel, width=800, height=12)
        self.bar.set(0)
        self.bar.grid(row=3, column=0, pady=10, padx=30)

        self.ctrl_btns = ctk.CTkFrame(self.bot_panel, fg_color="transparent")
        self.ctrl_btns.grid(row=4, column=0, pady=(10, 25))

        self.btn_go = ctk.CTkButton(self.ctrl_btns, text="🚀 INICIAR TRADUCCIÓN GLOBAL", command=self.start, 
                                    width=300, height=50, font=ctk.CTkFont(size=16, weight="bold"), state="disabled")
        self.btn_go.grid(row=0, column=0, padx=10)

        self.btn_stop = ctk.CTkButton(self.ctrl_btns, text="🛑 CANCELAR", command=self.stop, 
                                      width=150, height=50, fg_color="#c0392b", state="disabled")
        self.btn_stop.grid(row=0, column=1, padx=10)

    def add_folder(self):
        f = filedialog.askopenfilename(
            title="Navega hasta la carpeta y selecciona cualquier archivo dentro para usar esa carpeta",
            filetypes=[("Archivos Office", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt")]
        )
        
        if f:
            # Extraemos la ruta de la carpeta del archivo seleccionado
            folder_path = os.path.dirname(f)
            
            exts = ('.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt')
            files_found = False
            
            for file in os.listdir(folder_path):
                if file.lower().endswith(exts) and not file.startswith(('~$', 'BILINGUAL_')):
                    self.register_file(os.path.join(folder_path, file))
                    files_found = True
            
            if not files_found:
                messagebox.showwarning("Atención", "No se encontraron archivos compatibles en esta carpeta.")

    def add_file(self):
        f = filedialog.askopenfilename(filetypes=[("Archivos Office", "*.docx *.doc *.xlsx *.xls *.pptx *.ppt")])
        if f and not os.path.basename(f).startswith('BILINGUAL_'):
            self.register_file(f)

    def register_file(self, path):
        abs_p = os.path.abspath(path)
        if abs_p not in self.files_path_set:
            self.files_path_set.add(abs_p)
            
            row_frame = ctk.CTkFrame(self.list_frame, fg_color="transparent")
            row_frame.pack(fill="x", padx=5, pady=2)
            
            v = ctk.BooleanVar(value=True)
            chk = ctk.CTkCheckBox(row_frame, text=os.path.basename(path), variable=v)
            chk.pack(side="left", padx=10, fill="x", expand=True)
            
            btn_up = ctk.CTkButton(row_frame, text="▲", width=30, fg_color="#34495e", 
                                   command=lambda: self.move_file(abs_p, -1))
            btn_up.pack(side="right", padx=2)
            
            btn_down = ctk.CTkButton(row_frame, text="▼", width=30, fg_color="#34495e", 
                                     command=lambda: self.move_file(abs_p, 1))
            btn_down.pack(side="right", padx=2)
            
            self.files_to_process.append({
                "path": abs_p, 
                "var": v, 
                "frame": row_frame
            })
            self.btn_go.configure(state="normal")

    def clear(self):
        for w in self.list_frame.winfo_children(): w.destroy()
        self.files_to_process, self.files_path_set = [], set()
        self.btn_go.configure(state="disabled")

    def stop(self):
        self.is_cancelled = True
        self.lbl_status.configure(text="Cancelando...", text_color="#e74c3c")

    def start(self):
        sel = [item["path"] for item in self.files_to_process if item["var"].get()]
        if not sel: return

        self.is_cancelled = False
        self.btn_go.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.lbl_msg.configure(text=random.choice(romantic_notes))
        threading.Thread(target=self.main_loop, args=(sel,), daemon=True).start()

    def move_file(self, path, direction):
        idx = next(i for i, x in enumerate(self.files_to_process) if x["path"] == path)
        new_idx = idx + direction
        
        if 0 <= new_idx < len(self.files_to_process):
            self.files_to_process[idx], self.files_to_process[new_idx] = \
                self.files_to_process[new_idx], self.files_to_process[idx]
            
            self.refresh_file_list()

    def refresh_file_list(self):    
        for item in self.files_to_process:
            item["frame"].pack_forget()
        
        for item in self.files_to_process:
            item["frame"].pack(fill="x", padx=5, pady=2)

    def main_loop(self, files):
        self.lbl_status.configure(text="Analizando archivos y formatos legados...")
        self.total_units = 0
        self.done_units = 0
        ready_queue = []

        try:
            for p in files:
                if self.is_cancelled: break
                units, temp_p = self.prepare_document(p)
                self.total_units += units
                ready_queue.append((p, temp_p))

            if self.total_units == 0: self.total_units = 1
            self.start_time = time.time()

            for orig, temp in ready_queue:
                if self.is_cancelled: break
                self.lbl_status.configure(text=f"Procesando: {os.path.basename(orig)}", text_color="#ecf0f1")
                
                try:
                    if temp.endswith('.docx'):
                        self.translate_word(orig, temp)
                    elif temp.endswith('.xlsx'):
                        self.translate_excel(orig, temp)
                    elif temp.endswith('.pptx'):
                        self.translate_pptx(orig, temp)
                except Exception as e:
                    print(f"Error procesando {orig}: {e}")
                finally:
                    if "_TEMP_" in temp and os.path.exists(temp):
                        os.remove(temp)

        finally:
            self.bar.set(1 if not self.is_cancelled else 0)
            self.lbl_status.configure(text="¡Completado!" if not self.is_cancelled else "Cancelado", text_color="#2ecc71" if not self.is_cancelled else "#e74c3c")
            self.btn_go.configure(state="normal"); self.btn_stop.configure(state="disabled")
            self.after(0, self.show_love_note)

    def prepare_document(self, path):
        work_path = path
        if path.lower().endswith(('.doc', '.xls', '.ppt')) and win32:
            import pythoncom
            pythoncom.CoInitialize()
            app = None
            try:
                abs_path = os.path.abspath(path)
                if path.lower().endswith('.doc'):
                    app = win32.Dispatch('Word.Application')
                    doc = app.Documents.Open(abs_path)
                    work_path = abs_path + "_TEMP_.docx"
                    doc.SaveAs2(work_path, FileFormat=16)
                    doc.Close()
                elif path.lower().endswith('.xls'):
                    app = win32.Dispatch('Excel.Application')
                    app.DisplayAlerts = False
                    wb = app.Workbooks.Open(abs_path)
                    work_path = abs_path + "_TEMP_.xlsx"
                    wb.SaveAs(work_path, FileFormat=51)
                    wb.Close(False)
                elif path.lower().endswith('.ppt'):
                    app = win32.Dispatch('PowerPoint.Application')
                    pres = app.Presentations.Open(abs_path, WithWindow=False)
                    work_path = abs_path + "_TEMP_.pptx"
                    pres.SaveAs(work_path, FileFormat=24)
                    pres.Close()
            except Exception as e:
                print(f"Error conversión legacy: {e}")
            finally:
                if app:
                    app.Quit()
                    app = None
                pythoncom.CoUninitialize()
                self.cleanup_janitor()
                time.sleep(2.0)

        cnt = 0
        try:
            if work_path.lower().endswith('.docx'):
                d = Document(work_path)
                cnt = len(d.paragraphs) + sum(len(r.cells) for t in d.tables for r in t.rows)
            elif work_path.lower().endswith('.xlsx'):
                for _ in range(3):
                    try:
                        wb_count = openpyxl.load_workbook(work_path, read_only=True)
                        for s in wb_count.worksheets: 
                            cnt += (s.max_row if s.max_row else 1)
                        wb_count.close()
                        break
                    except:
                        time.sleep(1)
            elif work_path.lower().endswith('.pptx'):
                prs = Presentation(work_path)
                cnt = len(prs.slides)
        except: cnt = 1
        return cnt, work_path

    def translate_pptx(self, orig, temp):
        """Optimized PowerPoint Translation with Batching and Filtering."""
        try:
            prs = Presentation(temp)
            out_name = f"BILINGUAL_{os.path.basename(orig)}"
            if orig.lower().endswith('.ppt'): out_name += "x"
            output_path = os.path.join(os.path.dirname(orig), out_name)
            
            prs.save(output_path)
        except Exception as e:
            print(f"Error inicializando PPTX: {e}")
            return

        for slide_idx, slide in enumerate(prs.slides):
            if self.is_cancelled: return
            
            self.lbl_status.configure(text=f"Traduciendo Diapositiva {slide_idx + 1}/{len(prs.slides)}")
            
            for shape in slide.shapes:
                if self.is_cancelled: return
                
                if hasattr(shape, "text_frame") and shape.text.strip():
                    for paragraph in shape.text_frame.paragraphs:
                        clean_text = paragraph.text.strip()
                        if clean_text and len(clean_text) > 1:
                            try:
                                res = self.translator.translate(clean_text)
                                run = paragraph.add_run()
                                run.text = f"\n{res}"
                                if paragraph.runs:
                                    run.font.size = paragraph.runs[0].font.size
                            except Exception as e:
                                print(f"API Error in shape: {e}")
                                time.sleep(1)
                    self.upd_prog()
                
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            clean_cell = cell.text.strip()
                            if clean_cell:
                                try:
                                    res = self.translator.translate(clean_cell)
                                    cell.text = f"{clean_cell}\n{res}"
                                except: pass
                                self.upd_prog()

            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame
                if notes.text.strip():
                    try:
                        res = self.translator.translate(notes.text)
                        notes.text = f"{notes.text}\n---\n{res}"
                    except: pass

            try:
                prs.save(output_path)
            except: pass

    def translate_word(self, orig, temp):
        doc = Document(temp)
        out_name = f"BILINGUAL_{os.path.basename(orig)}"
        if orig.lower().endswith('.doc'): out_name += "x"
        output_path = os.path.join(os.path.dirname(orig), out_name)

        for p in doc.paragraphs:
            if self.is_cancelled: return
            if p.text.strip():
                try:
                    original_style = p.runs[0].font if p.runs else None
                    res = self.translator.translate(p.text)
                    
                    run = p.add_run(f"\n{res}")
                    run.italic = False
                    if original_style:
                        run.font.name = original_style.name
                        run.font.size = original_style.size
                    doc.save(output_path)
                except: pass
                self.upd_prog()
        
        for table in doc.tables:
            processed = set()
            for row in table.rows:
                for cell in row.cells:
                    if self.is_cancelled: return
                    if cell._tc in processed: continue
                    if cell.text.strip():
                        try:
                            original_text = cell.text
                            first_para = cell.paragraphs[0]
                            
                            res = self.translator.translate(original_text)
                            new_run = first_para.add_run(f"\n{res}")
                            new_run.italic = False
                            doc.save(output_path)
                            
                        except Exception as e:
                            print(f"Error en celda: {e}")
                        
                        self.upd_prog()
                    processed.add(cell._tc)

    def translate_excel(self, orig, temp):
        wb = openpyxl.load_workbook(temp, data_only=True)
        new_wb = openpyxl.Workbook()
        
        for idx, sn in enumerate(wb.sheetnames):
            if self.is_cancelled: break
            ws = wb[sn]
            nws = new_wb.active if idx == 0 else new_wb.create_sheet(sn)
            nws.title = sn

            for c, d in ws.column_dimensions.items(): nws.column_dimensions[c].width = d.width
            for r, d in ws.row_dimensions.items(): nws.row_dimensions[r].height = d.height

            max_row = ws.max_row
            max_col = ws.max_column
            
            for r in range(1, max_row + 1):
                for c in range(1, max_col + 1):
                    if self.is_cancelled: return
                    cell, ncell = ws.cell(row=r, column=c), nws.cell(row=r, column=c)
                    
                    if not isinstance(cell, MergedCell):
                        if isinstance(cell.value, str) and cell.value.strip():
                            try:
                                res = self.translator.translate(cell.value)
                                ncell.value = f"{cell.value}\n{res}"
                            except: ncell.value = cell.value
                            self.upd_prog()
                        else: ncell.value = cell.value

                    if cell.has_style:
                        ncell.font = copy.copy(cell.font)
                        ncell.border = copy.copy(cell.border)
                        ncell.fill = copy.copy(cell.fill)
                        ncell.alignment = Alignment(
                                wrap_text=True, 
                                horizontal=cell.alignment.horizontal or 'center', 
                                vertical=cell.alignment.vertical or 'center'
                            )

            for mr in ws.merged_cells.ranges: nws.merge_cells(str(mr))
            if hasattr(ws, '_images'):
                for img in list(ws._images):
                    try:
                        new_img = copy.copy(img)
                        nws.add_image(new_img)
                    except Exception as e:
                        print(f"No se pudo copiar una imagen: {e}")

        out_name = f"BILINGUAL_{os.path.basename(orig)}"
        if orig.lower().endswith('.xls'): out_name += "x"
        new_wb.save(os.path.join(os.path.dirname(orig), out_name))
        if "_TEMP_" in temp and os.path.exists(temp):
            try: os.remove(temp)
            except: pass

    def upd_prog(self):
        self.done_units += 1
        total = self.total_units if self.total_units > 0 else 1
        p = self.done_units / total
        self.bar.set(p)

        if random.random() < 0.10: 
            self.lbl_msg.configure(text=random.choice(romantic_notes))

        elapsed = time.time() - self.start_time
        if self.done_units > 0 and self.total_units > 0:
            avg_time = elapsed / self.done_units
            rem = avg_time * (self.total_units - self.done_units)
            m, s = divmod(int(max(0, rem)), 60)
            self.lbl_eta.configure(text=f"ETA Global: {m:02d}:{s:02d}")
        self.update_idletasks()
    
    def show_love_note(self):
        if self.is_cancelled:
            return

        note = ctk.CTkToplevel(self)
        note.title("Hecho con cariño")
        note.geometry("350x250")
        note.attributes("-topmost", True)
        
        selected_text = random.choice(final_love_notes)
        
        label = ctk.CTkLabel(
            note, 
            text=selected_text, 
            font=ctk.CTkFont(size=15, slant="italic"),
            justify="center"
        )
        label.pack(expand=True, padx=20, pady=20)
        
        btn = ctk.CTkButton(
            note, 
            text="Cerrar", 
            command=note.destroy, 
            fg_color="#db7093",
            hover_color="#c71585"
        )
        btn.pack(pady=(0, 20))

    def cleanup_janitor(self):
        processes = ["WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"]
        for proc in processes:
            try:
                subprocess.run(["taskkill", "/F", "/IM", proc, "/T"], 
                            stdout=subprocess.DEVNULL, 
                            stderr=subprocess.DEVNULL)
            except Exception:
                pass
        gc.collect()

if __name__ == "__main__":
    app = DocTranslatorPro()
    app.mainloop()