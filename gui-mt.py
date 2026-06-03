"""
GUI â€” MathType to LaTeX Converter + Quiz Pipeline
Tab 1: Cháº·ng 1 â€” Convert MathType OLE â†’ LaTeX text

"""

import sys
import os

# Fix Tcl/Tk initialization issues on Windows
if sys.platform == "win32":
    try:
        import tkinter as tk

        root = tk.Tk()
        root.destroy()
    except Exception:
        base_dir = None
        for p in sys.path:
            if p.endswith("Lib"):
                base_dir = os.path.dirname(p)
                break
            elif p.endswith("site-packages"):
                base_dir = os.path.dirname(os.path.dirname(p))
                break
        if base_dir:
            tcl_dir = os.path.join(base_dir, "tcl")
            if os.path.isdir(tcl_dir):
                for d in os.listdir(tcl_dir):
                    if d.startswith("tcl") and os.path.isdir(os.path.join(tcl_dir, d)):
                        os.environ["TCL_LIBRARY"] = os.path.join(tcl_dir, d)
                    elif d.startswith("tk") and os.path.isdir(os.path.join(tcl_dir, d)):
                        os.environ["TK_LIBRARY"] = os.path.join(tcl_dir, d)
                os.environ["PYTHONHOME"] = base_dir

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# Tab 1 â€” MathType â†’ LaTeX
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•


class ConvertTab:
    def __init__(self, parent):
        self.parent = parent  # the notebook tab frame
        self.root = parent.winfo_toplevel()

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.delete_var = tk.BooleanVar()
        self.delete_keyword = tk.StringVar()
        self.delete_mucdo_var = tk.BooleanVar()
        self.format_mcq_var = tk.BooleanVar()
        self.convert_tables_var = tk.BooleanVar()
        self.upload_images_var = tk.BooleanVar()  # NEW: Upload images to Cloudinary
        self.export_json_var = tk.BooleanVar()  # NEW: Option to export JSON

        self._build()

    def _build(self):
        pad = {"padx": 10, "pady": 4}
        f = self.parent

        # Input
        fr_in = ttk.Frame(f)
        fr_in.pack(fill="x", **pad)
        ttk.Label(fr_in, text="Input:").pack(side="left")
        ttk.Entry(fr_in, textvariable=self.input_path, width=55).pack(
            side="left", padx=4, fill="x", expand=True
        )
        ttk.Button(fr_in, text="Browse...", command=self._browse_input).pack(
            side="left"
        )

        # Output
        fr_out = ttk.Frame(f)
        fr_out.pack(fill="x", **pad)
        ttk.Label(fr_out, text="Output:").pack(side="left")
        ttk.Entry(fr_out, textvariable=self.output_path, width=55).pack(
            side="left", padx=4, fill="x", expand=True
        )
        ttk.Button(fr_out, text="Browse...", command=self._browse_output).pack(
            side="left"
        )

        # Options
        opt_frame = ttk.LabelFrame(f, text="Pre-processing options")
        opt_frame.pack(fill="x", padx=10, pady=2)

        fr_tables = ttk.Frame(opt_frame)
        fr_tables.pack(fill="x", padx=6, pady=2)
        ttk.Checkbutton(
            fr_tables,
            text="Chuyển bảng thành văn bản (Convert tables to text)",
            variable=self.convert_tables_var,
        ).pack(side="left")

        fr_del = ttk.Frame(opt_frame)
        fr_del.pack(fill="x", padx=6, pady=2)
        ttk.Checkbutton(
            fr_del, text="Xóa dòng chứa keyword:", variable=self.delete_var
        ).pack(side="left")
        ttk.Entry(fr_del, textvariable=self.delete_keyword, width=30).pack(
            side="left", padx=4, fill="x", expand=True
        )

        fr_mucdo = ttk.Frame(opt_frame)
        fr_mucdo.pack(fill="x", padx=6, pady=2)
        ttk.Checkbutton(
            fr_mucdo,
            text="Xóa mức độ ([Mức độ 1], [Mức độ 2], ...)",
            variable=self.delete_mucdo_var,
        ).pack(side="left")

        fr_mcq = ttk.Frame(opt_frame)
        fr_mcq.pack(fill="x", padx=6, pady=2)
        ttk.Checkbutton(
            fr_mcq,
            text="Format MCQ answers (B. C. D. each on new line)",
            variable=self.format_mcq_var,
        ).pack(side="left")

        fr_upload = ttk.Frame(opt_frame)
        fr_upload.pack(fill="x", padx=6, pady=2)
        ttk.Checkbutton(
            fr_upload,
            text="Upload ảnh lên Cloudinary (cần cấu hình cloudinary_config.txt)",
            variable=self.upload_images_var,
        ).pack(side="left")

        fr_json = ttk.Frame(opt_frame)
        fr_json.pack(fill="x", padx=6, pady=2)
        ttk.Checkbutton(
            fr_json,
            text="Chuyển nội dung sang file .json (như pipeline_cloudinary)",
            variable=self.export_json_var,
        ).pack(side="left")

        # Button + progress + log
        self.btn_convert = ttk.Button(f, text="Convert", command=self._start_convert)
        self.btn_convert.pack(pady=6)

        self.progress = ttk.Progressbar(f, mode="determinate", length=700)
        self.progress.pack(padx=10, pady=2, fill="x")

        self.log = scrolledtext.ScrolledText(
            f, height=14, state="disabled", font=("Consolas", 9)
        )
        self.log.pack(padx=10, pady=(4, 10), fill="both", expand=True)

    # File dialogs
    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Select Word file",
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)
            if not self.output_path.get():
                base, ext = os.path.splitext(path)
                self.output_path.set(base + "-latex" + ext)

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save output file",
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")
        self.root.update_idletasks()

    def _start_convert(self):
        inp = self.input_path.get().strip()
        out = self.output_path.get().strip()
        if not inp:
            messagebox.showwarning("Warning", "Please select an input file.")
            return
        if not os.path.isfile(inp):
            messagebox.showerror("Error", f"Input file not found:\n{inp}")
            return
        if not out:
            messagebox.showwarning("Warning", "Please specify an output file.")
            return
        self.btn_convert.configure(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.progress["value"] = 0
        threading.Thread(target=self._do_convert, args=(inp, out), daemon=True).start()

    def _do_convert(self, inp, out):
        try:
            self._log(f"Input:  {inp}")
            self._log(f"Output: {out}")
            self._log("=" * 60)
            self.root.after(0, lambda: self.progress.configure(mode="indeterminate"))
            self.root.after(0, self.progress.start)

            import io, contextlib, tempfile
            from convert_mathtype_to_latex import (
                process_docx,
                remove_paragraphs_containing,
                format_mcq_answers_in_docx,
                remove_level_tags_in_docx,
                convert_tables_to_text,
                format_all_text_times_new_roman_12,
            )

            buf = io.StringIO()
            tmp = inp
            temps = []

            if self.convert_tables_var.get():
                tmp2 = tempfile.mktemp(suffix=".docx")
                temps.append(tmp2)
                self._log("Chuyển bảng thành văn bản ...")
                with contextlib.redirect_stdout(buf):
                    convert_tables_to_text(tmp, tmp2)
                for line in buf.getvalue().splitlines():
                    self._log(line)
                buf = io.StringIO()
                tmp = tmp2

            if self.delete_var.get():
                kw = self.delete_keyword.get().strip()
                if kw:
                    tmp2 = tempfile.mktemp(suffix=".docx")
                    temps.append(tmp2)
                    self._log(f"Xóa dòng chứa: '{kw}' ...")
                    with contextlib.redirect_stdout(buf):
                        remove_paragraphs_containing(tmp, tmp2, kw)
                    for line in buf.getvalue().splitlines():
                        self._log(line)
                    buf = io.StringIO()
                    if tmp != inp and os.path.exists(tmp):
                        try:
                            os.remove(tmp)
                        except:
                            pass
                    tmp = tmp2

            if self.delete_mucdo_var.get():
                tmp2 = tempfile.mktemp(suffix=".docx")
                temps.append(tmp2)
                self._log("Xóa mức độ ...")
                with contextlib.redirect_stdout(buf):
                    remove_level_tags_in_docx(tmp, tmp2)
                for line in buf.getvalue().splitlines():
                    self._log(line)
                buf = io.StringIO()
                if tmp != inp and os.path.exists(tmp):
                    try:
                        os.remove(tmp)
                    except:
                        pass
                tmp = tmp2

            if self.format_mcq_var.get():
                tmp2 = tempfile.mktemp(suffix=".docx")
                temps.append(tmp2)
                self._log("Formatting MCQ answers ...")
                with contextlib.redirect_stdout(buf):
                    format_mcq_answers_in_docx(tmp, tmp2)
                for line in buf.getvalue().splitlines():
                    self._log(line)
                buf = io.StringIO()
                if tmp != inp and os.path.exists(tmp):
                    try:
                        os.remove(tmp)
                    except:
                        pass
                tmp = tmp2

            # Main conversion
            tmp2 = tempfile.mktemp(suffix=".docx")
            temps.append(tmp2)
            self._log("Chuyển đổi MathType sang LaTeX ...")
            with contextlib.redirect_stdout(buf):
                process_docx(tmp, tmp2)
            for line in buf.getvalue().splitlines():
                self._log(line)
            buf = io.StringIO()
            if tmp != inp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except:
                    pass
            tmp = tmp2

            # AUTOMATIC POST-PROCESSING:
            
            # Step 1: Remove [Mức độ 1], [Mức độ 2], [Mức độ 3], [Mức độ 4]
            tmp2 = tempfile.mktemp(suffix=".docx")
            temps.append(tmp2)
            self._log("Tự động xóa [Mức độ 1, 2, 3, 4] ...")
            with contextlib.redirect_stdout(buf):
                remove_level_tags_in_docx(tmp, tmp2)
            for line in buf.getvalue().splitlines():
                self._log(line)
            buf = io.StringIO()
            if tmp != inp and os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except:
                    pass
            tmp = tmp2

            # Step 2: Upload images to Cloudinary (AFTER MathType conversion)
            if self.upload_images_var.get():
                from convert_mathtype_to_latex import upload_images_to_cloudinary
                tmp2 = tempfile.mktemp(suffix=".docx")
                temps.append(tmp2)
                self._log("Upload ảnh lên Cloudinary (SAU khi convert LaTeX) ...")
                try:
                    with contextlib.redirect_stdout(buf):
                        upload_images_to_cloudinary(tmp, tmp2)
                    for line in buf.getvalue().splitlines():
                        self._log(line)
                    buf = io.StringIO()
                    if tmp != inp and os.path.exists(tmp):
                        try:
                            os.remove(tmp)
                        except:
                            pass
                    tmp = tmp2
                    self._log("✓ Upload ảnh hoàn tất")
                except Exception as e:
                    self._log(f"⚠ WARNING: Upload images failed: {e}")
                    self._log("Continuing without uploading images...")
                    # Don't fail the whole process, just skip this step
                    # Clean up temp file if upload failed
                    if os.path.exists(tmp2):
                        try:
                            os.remove(tmp2)
                        except:
                            pass

            # Step 3: Format all text with Times New Roman 12pt
            self._log("Tự động format Times New Roman 12pt ...")
            with contextlib.redirect_stdout(buf):
                format_all_text_times_new_roman_12(tmp, out)
            for line in buf.getvalue().splitlines():
                self._log(line)

            # Step 4: JSON Export
            if self.export_json_var.get():
                self._log("Bắt đầu chuyển nội dung sang file .json ...")
                try:
                    # Dynamically add D:\dethido or parent folders to path to import pipeline
                    dethido_path = "D:\\dethido"
                    if os.path.isdir(dethido_path) and dethido_path not in sys.path:
                        sys.path.insert(0, dethido_path)
                    
                    from pipeline.extractor import extract_questions
                    
                    self._log("Trích xuất câu hỏi từ Word ...")
                    questions = extract_questions(out)
                    self._log(f"Đã trích xuất {len(questions)} câu hỏi.")
                    
                    if questions:
                        # Setup Cloudinary config for image upload if needed
                        has_cloudinary = False
                        try:
                            # Try to configure from cloudinary_config.txt if env is not set
                            if not os.environ.get('CLOUDINARY_CLOUD_NAME'):
                                from upload_images_to_cloudinary import get_cloudinary_config
                                config = get_cloudinary_config()
                                if config.get('cloud_name'):
                                    os.environ['CLOUDINARY_CLOUD_NAME'] = config['cloud_name']
                                    os.environ['CLOUDINARY_API_KEY'] = config['api_key']
                                    os.environ['CLOUDINARY_API_SECRET'] = config['api_secret']
                                    os.environ['CLOUDINARY_FOLDER'] = config.get('folder', 'word_images')
                            
                            if os.environ.get('CLOUDINARY_CLOUD_NAME'):
                                has_cloudinary = True
                        except Exception as ce:
                            self._log(f"Lưu ý: Không thể cấu hình Cloudinary ({ce}). Sẽ xuất JSON không có ảnh.")

                        if has_cloudinary:
                            self._log("Upload ảnh lên Cloudinary cho JSON ...")
                            try:
                                from pipeline.image_uploader_cloudinary import extract_and_upload_images, replace_placeholders_with_urls
                                rel_to_url = extract_and_upload_images(out, questions)
                                questions = replace_placeholders_with_urls(questions, rel_to_url)
                                self._log(f"Đã upload {len(rel_to_url)} ảnh.")
                            except Exception as ue:
                                self._log(f"Cảnh báo: Không thể upload ảnh ({ue}). Tiếp tục tạo JSON ...")
                        
                        # Save local JSON backup matching the target schema
                        out_json_path = os.path.splitext(out)[0] + ".json"
                        mapped_questions = []
                        
                        diff_map = {
                            'Dễ': 'easy',
                            'Nhận biết': 'easy',
                            'Thông hiểu': 'easy',
                            'Trung bình': 'medium',
                            'Vận dụng': 'medium',
                            'Khó': 'hard',
                            'Vận dụng cao': 'hard'
                        }

                        for q in questions:
                            import re as _re
                            phan = q.get('phan', 'I')
                            raw_content = q.get('content', '')

                            # Remove [Image URL: ...] from raw_content
                            raw_content = _re.sub(r'\[Image\s*URL:\s*[^\]]+\]\s*', '', raw_content)

                            # ── Kiểm tra nội dung câu hỏi ──
                            _has_upper_options = bool(_re.search(
                                r'^\s*[A-D][.)][\s\S]',
                                raw_content, _re.MULTILINE
                            ))
                            _has_lower_options = bool(_re.search(
                                r'^\s*[a-d]\)[\s\S]',
                                raw_content, _re.MULTILINE
                            ))
                            _has_options = _has_upper_options or _has_lower_options

                            # Có "Lời giải" + "Đáp án:" trong nội dung
                            _has_loi_giai_dap_an = bool(_re.search(
                                r'Lời\s*giải\s*:?[\s\S]*?(?:Đáp\s*án|Đáp\s*số)\s*:',
                                raw_content, _re.IGNORECASE
                            ))

                            # Kiểm tra loidgiai có "Đáp án:" không
                            loidgiai_raw = q.get('loidgiai') or ''
                            loidgiai_raw = _re.sub(r'\[Image\s*URL:\s*[^\]]+\]\s*', '', loidgiai_raw)
                            _has_dap_an_in_loidgiai = bool(_re.search(
                                r'(?:Đáp\s*án|Đáp\s*số)\s*:',
                                loidgiai_raw, _re.IGNORECASE
                            ))

                            opts = q.get('options') or {}
                            # Clean [Image URL: ...] from options
                            opts = {k: _re.sub(r'\[Image\s*URL:\s*[^\]]+\]\s*', '', v) for k, v in opts.items()}

                            if phan == 'III':
                                q_type = 'sa'
                            elif phan == 'II':
                                q_type = 'msq'
                            elif phan == 'I':
                                has_actual_options = bool(opts)
                                if has_actual_options or _has_upper_options:
                                    q_type = 'mcq'
                                elif _has_lower_options:
                                    q_type = 'msq'
                                else:
                                    q_type = 'sa'
                            else:
                                has_actual_options = bool(opts)
                                if has_actual_options or _has_upper_options:
                                    q_type = 'mcq'
                                elif _has_lower_options:
                                    q_type = 'msq'
                                else:
                                    q_type = 'sa'

                            diff_str = q.get('metadata', {}).get('difficulty', 'Trung bình')
                            diff_level = diff_map.get(diff_str, 'medium')
                            
                            answer = q.get('answer') or ""
                            loidgiai = loidgiai_raw
                            content = raw_content
                            
                            # Phần III (tự luận): Xử lý đáp án
                            if q_type == 'sa':
                                import re

                                # Bước 1: Tách "Lời giải" ra khỏi content TRƯỚC
                                if re.search(r'Lời\s*giải\s*:?', content, re.IGNORECASE):
                                    parts = re.split(r'Lời\s*giải\s*:?\s*\n?', content, flags=re.IGNORECASE, maxsplit=1)
                                    content = parts[0].strip()
                                    if len(parts) > 1:
                                        loidgiai = (parts[1].strip() + "\n" + loidgiai).strip()
                                elif re.search(r'^Lời\s*giải\s*:?', loidgiai, re.IGNORECASE):
                                    loidgiai = re.sub(r'^Lời\s*giải\s*:?\s*\n?', '', loidgiai, flags=re.IGNORECASE).strip()

                                # Bước 2: Tìm đáp án trong content
                                m = re.search(r'(?:Đáp\s*án|Đáp\s*số)\s*:\s*\$?([0-9.,\-+/\\]+)\$?', content, re.IGNORECASE)
                                if m:
                                    answer = m.group(1).strip()
                                    content = re.sub(r'(?:Đáp\s*án|Đáp\s*số)\s*:\s*\$?[0-9.,\-+/\\]+\$?\s*\n?', '', content, flags=re.IGNORECASE).strip()

                                # Bước 3: Tìm đáp án trong loidgiai
                                m2 = re.search(r'(?:Đáp\s*án|Đáp\s*số)\s*:\s*\$?([0-9.,\-+/\\]+)\$?', loidgiai, re.IGNORECASE)
                                if m2:
                                    if not answer:
                                        answer = m2.group(1).strip()
                                    loidgiai = re.sub(r'(?:Đáp\s*án|Đáp\s*số)\s*:\s*\$?[0-9.,\-+/\\]+\$?\s*\n?', '', loidgiai, flags=re.IGNORECASE).strip()
                            
                            # Phần II: Lấy các ý ĐÚNG từ lời giải
                            elif q_type == 'msq':
                                if not answer and loidgiai:
                                    import re
                                    correct_items = []
                                    for line in loidgiai.split('\n'):
                                        m = re.match(r'^([a-d])\)\s+(?:ĐÚNG|Đúng)', line, re.IGNORECASE)
                                        if m:
                                            correct_items.append(m.group(1).upper())
                                    if correct_items:
                                        answer = ','.join(correct_items)

                            # Đổi ![](url) thành <img src="url" alt="de thi toan" />
                            import re
                            content = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", content)
                            
                            # Phần II: Cắt content tại dòng a) đầu tiên
                            if q_type == 'msq':
                                lines = content.split('\n')
                                question_lines = []
                                found_first_option = False
                                for line in lines:
                                    if re.match(r'^\s*[a-d]\)', line.strip()):
                                        found_first_option = True
                                        break
                                    question_lines.append(line)
                                content = '\n'.join(question_lines).strip()

                            explanation_html = loidgiai
                            explanation_html = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", explanation_html)
                            explanation_html = explanation_html.replace('\n', '<br>')
                            
                            mq = {
                                "type": q_type,
                                "question": content,
                                "correct_option": answer,
                                "explanation": explanation_html,
                                "difficulty_level": diff_level,
                                "is_dynamic": False
                            }

                            if q_type == 'mcq':
                                import re
                                opt_a = re.sub(r'^[Aa]\.\s*', '', opts.get('A', ''))
                                opt_b = re.sub(r'^[Bb]\.\s*', '', opts.get('B', ''))
                                opt_c = re.sub(r'^[Cc]\.\s*', '', opts.get('C', ''))
                                opt_d = re.sub(r'^[Dd]\.\s*', '', opts.get('D', ''))
                                
                                mq["option_a"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_a)
                                mq["option_b"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_b)
                                mq["option_c"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_c)
                                mq["option_d"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_d)
                            elif q_type == 'msq':
                                import re
                                full_content = q.get('content', '')
                                
                                items = {'a': '', 'b': '', 'c': '', 'd': ''}
                                lines = full_content.split('\n')
                                current_item = None
                                
                                for line in lines:
                                    m = re.match(r'^([a-d])\)\s*(.+)', line.strip())
                                    if m:
                                        current_item = m.group(1).lower()
                                        items[current_item] = m.group(2).strip()
                                    elif current_item and line.strip():
                                        items[current_item] += ' ' + line.strip()
                                
                                opt_a = re.sub(r'^[a-d]\)\s*', '', items.get('a', ''))
                                opt_b = re.sub(r'^[a-d]\)\s*', '', items.get('b', ''))
                                opt_c = re.sub(r'^[a-d]\)\s*', '', items.get('c', ''))
                                opt_d = re.sub(r'^[a-d]\)\s*', '', items.get('d', ''))
                                
                                mq["option_a"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_a)
                                mq["option_b"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_b)
                                mq["option_c"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_c)
                                mq["option_d"] = re.sub(r'!\[\]\(([^)]+)\)', r"<img src='\1' class='quiz-math-img' alt='de thi toan' />", opt_d)

                            mapped_questions.append(mq)

                        import json
                        from pathlib import Path
                        export_data = {
                            "title": Path(out).stem,
                            "questions": mapped_questions
                        }

                        with open(out_json_path, 'w', encoding='utf-8') as f:
                            json.dump(export_data, f, ensure_ascii=False, indent=2)
                        self._log(f"✓ Đã lưu file backup JSON: {out_json_path}")
                    else:
                        self._log("Cảnh báo: Không trích xuất được câu hỏi nào để lưu JSON.")
                except Exception as je:
                    self._log(f"⚠ LỖI khi xuất JSON: {je}")

            for t in temps:
                if os.path.exists(t):
                    try:
                        os.remove(t)
                    except:
                        pass

            self.root.after(0, self.progress.stop)
            self.root.after(
                0, lambda: self.progress.configure(mode="determinate", value=100)
            )
            self._log("\nDone!")
            self.root.after(0, lambda: self._show_success_dialog(out))

        except Exception as e:
            self.root.after(0, self.progress.stop)
            self.root.after(0, lambda: self.progress.configure(value=0))
            self._log(f"\nERROR: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.root.after(0, lambda: self.btn_convert.configure(state="normal"))

    def _show_success_dialog(self, out_path):
        dlg = tk.Toplevel(self.root)
        dlg.title("Conversion Complete")
        dlg.resizable(False, False)
        dlg.grab_set()
        fr_msg = ttk.Frame(dlg)
        fr_msg.pack(padx=20, pady=(18, 6))
        ttk.Label(
            fr_msg, text="Conversion complete!", font=("Segoe UI", 11, "bold")
        ).pack()
        ttk.Label(
            fr_msg,
            text=out_path,
            wraplength=480,
            foreground="#555",
            font=("Consolas", 8),
        ).pack(pady=(4, 0))
        fr_btn = ttk.Frame(dlg)
        fr_btn.pack(pady=(10, 16))

        def open_file():
            try:
                os.startfile(out_path)
            except:
                pass
            dlg.destroy()

        ttk.Button(fr_btn, text="Mở file Word", command=open_file, width=20).pack(
            side="left", padx=6
        )
        ttk.Button(fr_btn, text="Đóng", command=dlg.destroy, width=12).pack(
            side="left", padx=6
        )
        self.root.update_idletasks()
        rx = self.root.winfo_x() + self.root.winfo_width() // 2
        ry = self.root.winfo_y() + self.root.winfo_height() // 2
        dlg.update_idletasks()
        w, h = dlg.winfo_width(), dlg.winfo_height()
        dlg.geometry(f"+{rx - w // 2}+{ry - h // 2}")
        dlg.focus_set()


# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
# Main App
# â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Math Docx Tools")
        self.root.geometry("800x680")
        self.root.resizable(True, True)
        self._build()

    def _build(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)

        tab1 = ttk.Frame(nb)
        nb.add(tab1, text="Convert MathType to LaTeX ")
        self.convert_tab = ConvertTab(tab1)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
