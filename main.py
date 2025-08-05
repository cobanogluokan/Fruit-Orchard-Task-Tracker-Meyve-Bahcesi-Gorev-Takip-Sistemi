import pandas as pd
import os
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
from matplotlib.figure import Figure
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap.tooltip import ToolTip
import re

class FarmTaskTracker:
    """A class to manage farm task tracking with a GUI interface."""
    
    def __init__(self):
        # Set default directory for Excel file (user's home directory)
        self.data_dir = Path.home() / "FarmTasks"
        self.data_dir.mkdir(exist_ok=True)
        self.excel_file = self.data_dir / "farm_tasks.xlsx"
        self.selected_task_id = None
        self.root = None
        self.setup_excel_file()
        self.setup_gui()

    def setup_excel_file(self):
        """Create the Excel file if it doesn't exist."""
        try:
            if not self.excel_file.exists():
                df = pd.DataFrame(columns=[
                    "ID", "Job Name", "Description", "Start Date", "End Date",
                    "Estimated Cost", "Actual Cost", "Status"
                ])
                df.to_excel(self.excel_file, index=False)
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası oluşturulamadı: {e}")

    def load_tasks(self):
        """Load tasks from the Excel file."""
        try:
            return pd.read_excel(self.excel_file)
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası yüklenemedi: {e}")
            return pd.DataFrame()

    def validate_date(self, date_str):
        """Validate date format (YYYY-MM-DD)."""
        try:
            datetime.strptime(date_str, '%Y-%m-%d')
            return True
        except ValueError:
            return False

    def validate_cost(self, cost_str):
        """Validate cost input (numeric and non-negative)."""
        try:
            return float(cost_str) >= 0
        except (ValueError, TypeError):
            return False

    def add_task(self):
        """Add a new task to the Excel file."""
        try:
            name = self.entry_name.get().strip()
            desc = self.entry_desc.get().strip()
            start_date = self.entry_start.get().strip()
            end_date = self.entry_end.get().strip()
            estimated_cost = self.entry_estimated.get().strip()
            actual_cost = self.entry_actual.get().strip()
            status = "Done" if self.check_status.get() else "Waiting"

            # Input validation
            if not name:
                messagebox.showerror("Hata", "Görev adı boş olamaz!")
                return
            if not (self.validate_date(start_date) and self.validate_date(end_date)):
                messagebox.showerror("Hata", "Tarih formatı YYYY-MM-DD olmalı!")
                return
            if not self.validate_cost(estimated_cost):
                messagebox.showerror("Hata", "Tahmini maliyet geçerli bir sayı olmalı!")
                return
            if status == "Done" and not self.validate_cost(actual_cost):
                messagebox.showerror("Hata", "Gerçekleşen maliyet geçerli bir sayı olmalı!")
                return

            estimated_cost = float(estimated_cost)
            actual_cost = float(actual_cost) if status == "Done" else 0

            df = self.load_tasks()
            new_id = len(df) + 1
            new_task = pd.DataFrame([[new_id, name, desc, start_date, end_date,
                                    estimated_cost, actual_cost, status]], columns=df.columns)
            df = pd.concat([df, new_task], ignore_index=True)
            df.to_excel(self.excel_file, index=False)
            messagebox.showinfo("Başarılı", "Görev başarıyla eklendi!")
            self.clear_entries()
            self.show_tasks()
            self.status_label.configure(text="Görev eklendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Bir hata oluştu: {e}")

    def clear_entries(self):
        """Clear all input fields."""
        self.entry_name.delete(0, tk.END)
        self.entry_desc.delete(0, tk.END)
        self.entry_start.delete(0, tk.END)
        self.entry_end.delete(0, tk.END)
        self.entry_estimated.delete(0, tk.END)
        self.entry_actual.delete(0, tk.END)
        self.check_status.set(0)

    def load_selected_task(self):
        """Load selected task into input fields."""
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("Uyarı", "Lütfen düzenlemek için bir görev seçin.")
            return

        values = self.table.item(selected, "values")
        self.selected_task_id = int(values[0])

        self.clear_entries()
        self.entry_name.insert(0, values[1])
        self.entry_desc.insert(0, values[2])
        self.entry_start.insert(0, values[3])
        self.entry_end.insert(0, values[4])
        self.entry_estimated.insert(0, values[5])
        self.entry_actual.insert(0, values[6])
        self.check_status.set(1 if values[7] == "Done" else 0)
        self.status_label.configure(text="Görev düzenlenmek için yüklendi.")

    def update_task(self):
        """Update the selected task."""
        if self.selected_task_id is None:
            messagebox.showerror("Hata", "Güncelleme için görev seçilmedi.")
            return
        try:
            df = self.load_tasks()
            idx = df[df["ID"] == self.selected_task_id].index[0]

            name = self.entry_name.get().strip()
            desc = self.entry_desc.get().strip()
            start_date = self.entry_start.get().strip()
            end_date = self.entry_end.get().strip()
            estimated_cost = self.entry_estimated.get().strip()
            actual_cost = self.entry_actual.get().strip()
            status = "Done" if self.check_status.get() else "Waiting"

            # Input validation
            if not name:
                messagebox.showerror("Hata", "Görev adı boş olamaz!")
                return
            if not (self.validate_date(start_date) and self.validate_date(end_date)):
                messagebox.showerror("Hata", "Tarih formatı YYYY-MM-DD olmalı!")
                return
            if not self.validate_cost(estimated_cost):
                messagebox.showerror("Hata", "Tahmini maliyet geçerli bir sayı olmalı!")
                return
            if status == "Done" and not self.validate_cost(actual_cost):
                messagebox.showerror("Hata", "Gerçekleşen maliyet geçerli bir sayı olmalı!")
                return

            df.at[idx, "Job Name"] = name
            df.at[idx, "Description"] = desc
            df.at[idx, "Start Date"] = start_date
            df.at[idx, "End Date"] = end_date
            df.at[idx, "Estimated Cost"] = float(estimated_cost)
            df.at[idx, "Actual Cost"] = float(actual_cost) if status == "Done" else 0
            df.at[idx, "Status"] = status

            df.to_excel(self.excel_file, index=False)
            messagebox.showinfo("Başarılı", "Görev başarıyla güncellendi.")
            self.selected_task_id = None
            self.clear_entries()
            self.show_tasks()
            self.status_label.configure(text="Görev güncellendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Güncelleme başarısız: {e}")

    def delete_task(self):
        """Delete the selected task."""
        try:
            selected_item = self.table.selection()
            if not selected_item:
                messagebox.showerror("Hata", "Lütfen silmek için bir görev seçin!")
                return

            result = messagebox.askyesno("Onay", "Bu görevi silmek istediğinizden emin misiniz?")
            if not result:
                return

            df = self.load_tasks()
            selected_id = int(self.table.item(selected_item, "values")[0])
            df = df[df["ID"] != selected_id]
            df["ID"] = range(1, len(df) + 1)
            df.to_excel(self.excel_file, index=False)
            messagebox.showinfo("Başarılı", "Görev başarıyla silindi!")
            self.clear_entries()
            self.show_tasks()
            self.status_label.configure(text="Görev silindi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Bir hata oluştu: {e}")

    def show_tasks(self, filter_status=None):
        """Display tasks in the table, optionally filtered by status."""
        try:
            df = self.load_tasks()
            if filter_status:
                df = df[df["Status"] == filter_status]
            self.table.delete(*self.table.get_children())
            for _, row in df.iterrows():
                self.table.insert("", "end", values=row.tolist())
            self.status_label.configure(text=f"{len(df)} görev görüntülendi.")
        except Exception as e:
            messagebox.showerror("Hata", f"Görevler yüklenemedi: {e}")

    def show_calendar(self, filtered_df=None, save_as_pdf=False):
        """Display a calendar view of tasks."""
        try:
            df = filtered_df if filtered_df is not None else self.load_tasks()
            df = df.copy()
            df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
            df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")
            df = df.dropna(subset=["Start Date", "End Date"])

            grouped = df.groupby("Job Name")
            collapsed_tasks = []

            for name, group in grouped:
                start = group["Start Date"].min()
                end = group["End Date"].max()
                desc = " / ".join(group["Description"].dropna().astype(str).unique())
                collapsed_tasks.append({
                    "Job Name": name,
                    "Start Date": start,
                    "End Date": end,
                    "Descriptions": desc,
                    "Details": group
                })

            collapsed_tasks.sort(key=lambda x: x["Start Date"])
            fig, ax = plt.subplots(figsize=(14, max(6, len(collapsed_tasks) * 0.5)))

            total_actual = 0
            total_estimated = 0

            for idx, task in enumerate(collapsed_tasks):
                start = task["Start Date"]
                end = task["End Date"]
                duration = (end - start).days or 1

                done_cost = task["Details"].loc[task["Details"]["Status"] == "Done", "Actual Cost"].sum()
                wait_cost = task["Details"].loc[task["Details"]["Status"] == "Waiting", "Estimated Cost"].sum()
                total_cost = done_cost + wait_cost

                done_ratio = done_cost / total_cost if total_cost else 0
                wait_ratio = wait_cost / total_cost if total_cost else 0

                done_days = int(duration * done_ratio)
                wait_days = int(duration * wait_ratio)

                done_end = start + pd.Timedelta(days=done_days)
                wait_end = done_end + pd.Timedelta(days=wait_days)

                if done_cost > 0:
                    ax.plot([start, done_end], [idx, idx], color="green", linewidth=4)
                    ax.text(done_end, idx + 0.1, f"Harcanan: {done_cost:,.0f} EUR".replace(",", "."), fontsize=8, va="bottom")
                    total_actual += done_cost

                if wait_cost > 0:
                    ax.plot([done_end, wait_end], [idx, idx], color="orange", linewidth=4)
                    ax.text(wait_end, idx + 0.1, f"Beklenen: {wait_cost:,.0f} EUR".replace(",", "."), fontsize=8, va="bottom")
                    total_estimated += wait_cost

            ax.set_yticks(range(len(collapsed_tasks)))
            ax.set_yticklabels([t["Job Name"] for t in collapsed_tasks], fontsize=9)
            ax.set_title("Görev Zaman Çizelgesi ve Maliyet", fontsize=12, pad=10)
            ax.grid(True, linestyle="--", linewidth=0.5)
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
            ax.xaxis.set_major_locator(mdates.MonthLocator())
            plt.xticks(rotation=45)
            plt.tight_layout()

            ax.text(
                0.01, -0.12,
                f"Harcanan: {total_actual:,.0f} EUR | Beklenen: {total_estimated:,.0f} EUR | Toplam: {(total_actual + total_estimated):,.0f} EUR".replace(",", "."),
                transform=ax.transAxes,
                fontsize=10,
                color='black',
                ha='left'
            )

            if save_as_pdf:
                file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
                if file_path:
                    fig.savefig(file_path, format="pdf", bbox_inches='tight')
                    messagebox.showinfo("Başarılı", f"Takvim PDF olarak kaydedildi: {file_path}")
                    self.status_label.configure(text="Takvim PDF olarak kaydedildi.")
                plt.close(fig)
            else:
                cal_window = ttk.Toplevel(self.root)
                cal_window.title("Takvim Görünümü")
                cal_window.geometry("1000x600")
                canvas = FigureCanvasTkAgg(fig, master=cal_window)
                canvas.draw()
                canvas.get_tk_widget().pack(fill="both", expand=True)
                ttk.Button(cal_window, text="PDF Olarak Kaydet", command=lambda: self.show_calendar(filtered_df, save_as_pdf=True), style="primary.TButton").pack(pady=10)
                self.status_label.configure(text="Takvim açıldı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Takvim oluşturulurken hata: {e}")

    def export_calendar_pdf(self):
        """Export calendar as PDF."""
        self.show_calendar(save_as_pdf=True)

    def open_calendar_selection(self):
        """Open a window to select date range for expense calculation and calendar display."""
        try:
            popup = ttk.Toplevel(self.root)
            popup.title("Tarih Aralığı Seç")
            popup.geometry("300x300")

            ttk.Label(popup, text="Başlangıç Tarihi:").pack(pady=5)
            cal_start = Calendar(popup, selectmode='day', date_pattern='yyyy-mm-dd')
            cal_start.pack(pady=5)

            ttk.Label(popup, text="Bitiş Tarihi:").pack(pady=5)
            cal_end = Calendar(popup, selectmode='day', date_pattern='yyyy-mm-dd')
            cal_end.pack(pady=5)

            def calculate_and_show_calendar():
                try:
                    start = pd.to_datetime(cal_start.get_date())
                    end = pd.to_datetime(cal_end.get_date())

                    if start > end:
                        messagebox.showerror("Hata", "Başlangıç tarihi bitiş tarihinden sonra olamaz!")
                        return

                    df = self.load_tasks()
                    df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
                    df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")

                    mask = (df["Start Date"] >= start) & (df["End Date"] <= end)
                    filtered = df.loc[mask].copy()

                    if filtered.empty:
                        messagebox.showinfo("Bilgi", "Belirtilen tarihler arasında görev bulunamadı.")
                        return

                    total_actual = filtered.loc[filtered["Status"] == "Done", "Actual Cost"].sum()
                    total_estimated = filtered.loc[filtered["Status"] == "Waiting", "Estimated Cost"].sum()

                    result_text = f"Toplam Gerçekleşen Maliyet: {total_actual:,.2f} EUR\nToplam Tahmini Maliyet: {total_estimated:,.2f} EUR"
                    messagebox.showinfo("Harcamalar", result_text)
                    self.show_calendar(filtered_df=filtered)
                    self.status_label.configure(text="Harcama hesaplandı ve takvim gösterildi.")
                except Exception as e:
                    messagebox.showerror("Hata", f"Bir hata oluştu: {e}")

            ttk.Button(popup, text="Hesapla ve Göster", command=calculate_and_show_calendar, style="primary.TButton").pack(pady=10)
        except Exception as e:
            messagebox.showerror("Hata", f"Tarih seçimi penceresi açılırken hata: {e}")

    def export_all_tasks(self):
        """Export all tasks to a new Excel file."""
        try:
            df = self.load_tasks()
            if df.empty:
                messagebox.showinfo("Bilgi", "Aktarılacak görev bulunamadı.")
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )

            if not file_path:
                return

            now = datetime.now().strftime("%Y-%m-%d")
            writer = pd.ExcelWriter(file_path, engine='openpyxl')

            df.to_excel(writer, index=False, sheet_name='Tüm Görevler')

            summary = pd.DataFrame({
                'Toplam Görev Sayısı': [len(df)],
                'Tamamlanan Görev Sayısı': [len(df[df['Status'] == 'Done'])],
                'Bekleyen Görev Sayısı': [len(df[df['Status'] == 'Waiting'])],
                'Toplam Gerçekleşen Maliyet': [df[df['Status'] == 'Done']['Actual Cost'].sum()],
                'Toplam Beklenen Maliyet': [df[df['Status'] == 'Waiting']['Estimated Cost'].sum()],
                'Rapor Tarihi': [now]
            })

            summary.to_excel(writer, index=False, sheet_name='Özet')
            writer.close()
            messagebox.showinfo("Başarılı", f"Tüm görevler dışa aktarıldı: {file_path}")
            self.status_label.configure(text="Görevler Excel'e aktarıldı.")
        except Exception as e:
            messagebox.showerror("Hata", f"Dışa aktarma sırasında hata: {e}")

    def show_statistics(self):
        """Display task statistics."""
        try:
            df = self.load_tasks()
            if df.empty:
                messagebox.showinfo("Bilgi", "İstatistik gösterilecek görev bulunamadı.")
                return

            df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
            df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")

            stats_window = ttk.Toplevel(self.root)
            stats_window.title("Görev İstatistikleri")
            stats_window.geometry("800x600")

            total_tasks = len(df)
            done_tasks = len(df[df["Status"] == "Done"])
            waiting_tasks = len(df[df["Status"] == "Waiting"])
            total_actual_cost = df[df["Status"] == "Done"]["Actual Cost"].sum()
            total_estimated_cost = df[df["Status"] == "Waiting"]["Estimated Cost"].sum()
            df["Duration"] = (df["End Date"] - df["Start Date"]).dt.days
            avg_duration = df["Duration"].mean()

            info_frame = ttk.Frame(stats_window)
            info_frame.pack(pady=10, fill="x")
            ttk.Label(info_frame, text=f"Toplam Görev Sayısı: {total_tasks}", font=("Arial", 12)).pack(anchor="w", pady=2)
            ttk.Label(info_frame, text=f"Tamamlanan Görev Sayısı: {done_tasks}", font=("Arial", 12)).pack(anchor="w", pady=2)
            ttk.Label(info_frame, text=f"Bekleyen Görev Sayısı: {waiting_tasks}", font=("Arial", 12)).pack(anchor="w", pady=2)
            ttk.Label(info_frame, text=f"Toplam Gerçekleşen Maliyet: {total_actual_cost:,.2f} EUR", font=("Arial", 12)).pack(anchor="w", pady=2)
            ttk.Label(info_frame, text=f"Toplam Beklenen Maliyet: {total_estimated_cost:,.2f} EUR", font=("Arial", 12)).pack(anchor="w", pady=2)
            ttk.Label(info_frame, text=f"Ortalama Görev Süresi: {avg_duration:.1f} gün", font=("Arial", 12)).pack(anchor="w", pady=2)

            fig = Figure(figsize=(8, 8))
            ax1 = fig.add_subplot(221)
            status_counts = df["Status"].value_counts()
            ax1.pie(
                status_counts,
                labels=status_counts.index,
                autopct='%1.1f%%',
                colors=['green', 'orange'] if "Done" in status_counts.index and "Waiting" in status_counts.index else None
            )
            ax1.set_title("Görev Durumu Dağılımı")

            ax2 = fig.add_subplot(222)
            cost_data = [total_actual_cost, total_estimated_cost]
            ax2.bar(['Gerçekleşen', 'Beklenen'], cost_data, color=['green', 'orange'])
            ax2.set_title("Maliyet Karşılaştırma")
            ax2.set_ylabel("EUR")

            ax3 = fig.add_subplot(212)
            df['Month'] = df['Start Date'].dt.to_period('M')
            monthly_tasks = df.groupby('Month').size()
            ax3.plot(range(len(monthly_tasks)), monthly_tasks.values, marker='o')
            ax3.set_xticks(range(len(monthly_tasks)))
            ax3.set_xticklabels([str(period) for period in monthly_tasks.index], rotation=45)
            ax3.set_title("Aylık Görev Sayısı")
            ax3.set_ylabel("Görev Sayısı")
            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=stats_window)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            self.status_label.configure(text="İstatistikler gösterildi.")
        except Exception as e:
            messagebox.showerror("Hata", f"İstatistikler oluşturulurken hata: {e}")

    def select_date(self, entry_widget):
        """Show a date picker and set the selected date in the entry field."""
        def set_date():
            selected_date = cal.get_date()
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, selected_date)
            date_popup.destroy()

        date_popup = ttk.Toplevel(self.root)
        date_popup.title("Tarih Seç")
        date_popup.geometry("300x250")
        cal = Calendar(date_popup, selectmode='day', date_pattern='yyyy-mm-dd')
        cal.pack(pady=10)
        ttk.Button(date_popup, text="Seç", command=set_date, style="primary.TButton").pack(pady=5)

    def setup_gui(self):
        """Set up the GUI with ttkbootstrap."""
        self.root = ttk.Window(themename="flatly")  # Modern theme
        self.root.title("Çiftlik Görev Takip")
        self.root.geometry("1000x800")

        # Main frames
        input_frame = ttk.LabelFrame(self.root, text="Görev Bilgileri", padding=10)
        input_frame.pack(fill="x", padx=10, pady=5)
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill="x", padx=10, pady=5)
        table_frame = ttk.Frame(self.root)
        table_frame.pack(fill="both", expand=True, padx=10, pady=5)
        status_frame = ttk.Frame(self.root)
        status_frame.pack(fill="x", padx=10, pady=5)

        # Status bar
        self.status_label = ttk.Label(status_frame, text="Hazır", relief="sunken", anchor="w")
        self.status_label.pack(fill="x")

        # Input fields
        field_width = 40
        row1 = ttk.Frame(input_frame)
        row1.pack(fill="x", padx=5, pady=5)
        ttk.Label(row1, text="Görev Adı:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.entry_name = ttk.Entry(row1, width=field_width)
        self.entry_name.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ToolTip(self.entry_name, "Görev adını buraya girin (örn: Elma Hasadı)")
        ttk.Label(row1, text="Açıklama:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
        self.entry_desc = ttk.Entry(row1, width=field_width)
        self.entry_desc.grid(row=0, column=3, padx=5, pady=2, sticky="w")
        ToolTip(self.entry_desc, "Görevle ilgili detaylı açıklama")

        row2 = ttk.Frame(input_frame)
        row2.pack(fill="x", padx=5, pady=5)
        ttk.Label(row2, text="Başlangıç Tarihi:").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        date_frame1 = ttk.Frame(row2)
        date_frame1.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        self.entry_start = ttk.Entry(date_frame1, width=field_width-5)
        self.entry_start.pack(side="left")
        ttk.Button(date_frame1, text="...", width=3, command=lambda: self.select_date(self.entry_start), style="secondary.TButton").pack(side="left")
        ToolTip(self.entry_start, "Başlangıç tarihi (YYYY-MM-DD)")
        ttk.Label(row2, text="Bitiş Tarihi:").grid(row=0, column=2, padx=5, pady=2, sticky="w")
        date_frame2 = ttk.Frame(row2)
        date_frame2.grid(row=0, column=3, padx=5, pady=2, sticky="w")
        self.entry_end = ttk.Entry(date_frame2, width=field_width-5)
        self.entry_end.pack(side="left")
        ttk.Button(date_frame2, text="...", width=3, command=lambda: self.select_date(self.entry_end), style="secondary.TButton").pack(side="left")
        ToolTip(self.entry_end, "Bitiş tarihi (YYYY-MM-DD)")

        row3 = ttk.Frame(input_frame)
        row3.pack(fill="x", padx=5, pady=5)
        ttk.Label(row3, text="Tahmini Maliyet (EUR):").grid(row=0, column=0, padx=5, pady=2, sticky="w")
        self.entry_estimated = ttk.Entry(row3, width=field_width)
        self.entry_estimated.grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ToolTip(self.entry_estimated, "Tahmini maliyet (sayı, örn: 1000)")
        self.check_status = tk.BooleanVar()
        status_check = ttk.Checkbutton(row3, text="Tamamlandı?", variable=self.check_status)
        status_check.grid(row=0, column=2, padx=5, pady=2, sticky="w")
        ToolTip(status_check, "Görevin tamamlanıp tamamlanmadığını işaretleyin")
        ttk.Label(row3, text="Gerçekleşen Maliyet (EUR):").grid(row=0, column=3, padx=5, pady=2, sticky="w")
        self.entry_actual = ttk.Entry(row3, width=field_width-10)
        self.entry_actual.grid(row=0, column=4, padx=5, pady=2, sticky="w")
        ToolTip(self.entry_actual, "Gerçekleşen maliyet (tamamlandıysa, örn: 950)")

        # Buttons
        button_width = 20
        button_row1 = ttk.Frame(button_frame)
        button_row1.pack(fill="x", padx=5, pady=2)
        ttk.Button(button_row1, text="Görev Ekle", width=button_width, command=self.add_task, style="primary.TButton").pack(side="left", padx=5)
        ttk.Button(button_row1, text="Görevi Düzenle", width=button_width, command=self.load_selected_task, style="secondary.TButton").pack(side="left", padx=5)
        ttk.Button(button_row1, text="Görevi Güncelle", width=button_width, command=self.update_task, style="success.TButton").pack(side="left", padx=5)
        ttk.Button(button_row1, text="Görevi Sil", width=button_width, command=self.delete_task, style="danger.TButton").pack(side="left", padx=5)

        button_row2 = ttk.Frame(button_frame)
        button_row2.pack(fill="x", padx=5, pady=2)
        ttk.Button(button_row2, text="Tüm Görevler", width=button_width, command=lambda: self.show_tasks(None), style="info.TButton").pack(side="left", padx=5)
        ttk.Button(button_row2, text="Tamamlananlar", width=button_width, command=lambda: self.show_tasks("Done"), style="success.TButton").pack(side="left", padx=5)
        ttk.Button(button_row2, text="Bekleyenler", width=button_width, command=lambda: self.show_tasks("Waiting"), style="warning.TButton").pack(side="left", padx=5)

        button_row3 = ttk.Frame(button_frame)
        button_row3.pack(fill="x", padx=5, pady=2)
        ttk.Button(button_row3, text="Takvimi Aç", width=button_width, command=self.show_calendar, style="primary.TButton").pack(side="left", padx=5)
        ttk.Button(button_row3, text="Takvimi PDF'e Aktar", width=button_width, command=self.export_calendar_pdf, style="secondary.TButton").pack(side="left", padx=5)
        ttk.Button(button_row3, text="Harcama Hesapla", width=button_width, command=self.open_calendar_selection, style="info.TButton").pack(side="left", padx=5)
        ttk.Button(button_row3, text="İstatistikleri Göster", width=button_width, command=self.show_statistics, style="success.TButton").pack(side="left", padx=5)

        button_row4 = ttk.Frame(button_frame)
        button_row4.pack(fill="x", padx=5, pady=2)
        ttk.Button(button_row4, text="Tüm Görevleri Dışa Aktar", width=button_width, command=self.export_all_tasks, style="primary.TButton").pack(side="left", padx=5)
        ttk.Button(button_row4, text="Temizle", width=button_width, command=self.clear_entries, style="secondary.TButton").pack(side="left", padx=5)

        # Table
        columns = ("ID", "Job Name", "Description", "Start Date", "End Date", "Estimated Cost", "Actual Cost", "Status")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", style="Treeview")
        column_widths = {
            "ID": 50,
            "Job Name": 150,
            "Description": 200,
            "Start Date": 100,
            "End Date": 100,
            "Estimated Cost": 100,
            "Actual Cost": 100,
            "Status": 80
        }
        for col in columns:
            self.table.heading(col, text=col)
            self.table.column(col, width=column_widths[col])
        scrollbar_y = ttk.Scrollbar(table_frame, orient="vertical", command=self.table.yview)
        scrollbar_x = ttk.Scrollbar(table_frame, orient="horizontal", command=self.table.xview)
        self.table.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.table.pack(side="left", fill="both", expand=True)

        # Show initial tasks
        self.show_tasks()

    def run(self):
        """Start the application."""
        self.root.mainloop()

if __name__ == "__main__":
    try:
        import pandas
        import openpyxl
        import tkcalendar
        import matplotlib
    except ImportError as e:
        print(f"Hata: Gerekli kütüphane eksik: {e}")
        print("Lütfen aşağıdaki komutları çalıştırarak gerekli kütüphaneleri kurun:")
        print("pip install pandas openpyxl tkcalendar matplotlib ttkbootstrap")
        exit(1)

    app = FarmTaskTracker()
    app.run()
