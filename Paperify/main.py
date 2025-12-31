import sys
import csv
import textwrap
import re
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QPushButton, QComboBox,
    QListWidget, QMessageBox, QLineEdit, QSpinBox, 
    QFormLayout, QGroupBox, QScrollArea, QFileDialog,
    QDialog, QDialogButtonBox, QFrame
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor

# --- MATPLOTLIB IMPORTS ---
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch
import matplotlib.font_manager as fm

# --- URDU TEXT HANDLERS ---
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_URDU_LIB = True
except ImportError:
    HAS_URDU_LIB = False
    print("Warning: 'arabic-reshaper' or 'python-bidi' not installed. Urdu will look broken.")

# --- DEFAULT FONT SETTINGS ---
plt.rcParams['font.family'] = 'serif'
plt.rcParams['mathtext.fontset'] = 'cm' 
plt.rcParams['axes.unicode_minus'] = False 

# =============================================================================
#    DIALOG: STARTUP DETAILS
# =============================================================================
class SetupDetailsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Exam Details - Dar-e-Arqam School")
        self.setModal(True)
        self.resize(450, 350)
        self.data = {}
        
        layout = QVBoxLayout()
        
        # Branding Header
        lbl_brand = QLabel("DAR-E-ARQAM SCHOOL")
        lbl_brand.setAlignment(Qt.AlignCenter)
        lbl_brand.setStyleSheet("font-size: 20px; font-weight: bold; color: #2c3e50; margin-bottom: 5px;")
        layout.addWidget(lbl_brand)
        
        lbl_sub = QLabel("Exam Generator Setup")
        lbl_sub.setAlignment(Qt.AlignCenter)
        lbl_sub.setStyleSheet("font-size: 12px; color: #7f8c8d; margin-bottom: 15px;")
        layout.addWidget(lbl_sub)
        
        form_layout = QFormLayout()
        form_layout.setLabelAlignment(Qt.AlignRight)
        
        self.inp_school = QLineEdit("DAR-E-ARQAM SCHOOL")
        self.inp_test = QLineEdit("Revision Test")
        self.inp_class = QLineEdit()
        self.inp_class.setPlaceholderText("One")
        self.inp_subject = QLineEdit()
        self.inp_subject.setPlaceholderText("Urdu")
        self.inp_time = QLineEdit("1 Hour 30 Mins")
        self.inp_marks = QLineEdit("50")
        
        # Style inputs
        for inp in [self.inp_school, self.inp_test, self.inp_class, self.inp_subject, self.inp_time, self.inp_marks]:
            inp.setStyleSheet("padding: 6px; border: 1px solid #ccc; border-radius: 4px;")

        form_layout.addRow("School Name:", self.inp_school)
        form_layout.addRow("Test Title:", self.inp_test)
        form_layout.addRow("Class:", self.inp_class)
        form_layout.addRow("Subject:", self.inp_subject)
        form_layout.addRow("Total Time:", self.inp_time)
        form_layout.addRow("Total Marks:", self.inp_marks)
        
        layout.addLayout(form_layout)
        
        # Buttons
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.validate_and_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        
        self.setLayout(layout)

    def validate_and_accept(self):
        if not self.inp_class.text() or not self.inp_subject.text():
            QMessageBox.warning(self, "Missing Info", "Please enter Class and Subject.")
            return
        
        self.data = {
            "school": self.inp_school.text(),
            "test": self.inp_test.text(),
            "class": self.inp_class.text(),
            "subject": self.inp_subject.text(),
            "time": self.inp_time.text(),
            "marks": self.inp_marks.text()
        }
        self.accept()

# =============================================================================
#    DIALOG: NEW SECTION SETUP (POP-UP)
# =============================================================================
class SectionSetupDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Section")
        self.setModal(True)
        self.resize(400, 300)
        self.section_data = None

        layout = QVBoxLayout()
        layout.addWidget(QLabel("<b>Section Configuration</b>"))

        form = QFormLayout()
        
        self.inp_title = QLineEdit()
        self.inp_title.setPlaceholderText("e.g. Question 2")
        
        self.inp_desc = QLineEdit()
        self.inp_desc.setPlaceholderText("e.g. Write short answers (Any 6)")

        # Marks Logic
        self.spin_marks_per_q = QSpinBox()
        self.spin_marks_per_q.setRange(1, 100)
        self.spin_marks_per_q.setValue(2)

        self.spin_attempt_count = QSpinBox()
        self.spin_attempt_count.setRange(1, 100)
        self.spin_attempt_count.setValue(5)

        self.lbl_total_calc = QLabel("Total: 10 Marks")
        self.lbl_total_calc.setStyleSheet("color: #27ae60; font-weight: bold;")

        # Update total label when spins change
        self.spin_marks_per_q.valueChanged.connect(self.update_total)
        self.spin_attempt_count.valueChanged.connect(self.update_total)

        form.addRow("Title:", self.inp_title)
        form.addRow("Description:", self.inp_desc)
        form.addRow("Marks per Question:", self.spin_marks_per_q)
        form.addRow("Questions to Attempt:", self.spin_attempt_count)
        form.addRow("Calculated Total:", self.lbl_total_calc)

        layout.addLayout(form)
        
        layout.addWidget(QLabel("<small><i>Note: You can add as many questions as you want later.<br>This only sets the marking scheme.</i></small>"))

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.save_data)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def update_total(self):
        total = self.spin_marks_per_q.value() * self.spin_attempt_count.value()
        self.lbl_total_calc.setText(f"Total: {total} Marks")

    def save_data(self):
        if not self.inp_title.text():
            QMessageBox.warning(self, "Error", "Section Title is required.")
            return
        
        self.section_data = {
            "name": self.inp_title.text(),
            "desc": self.inp_desc.text(),
            "marks_per_q": self.spin_marks_per_q.value(),
            "attempt_count": self.spin_attempt_count.value(),
            "total_marks": self.spin_marks_per_q.value() * self.spin_attempt_count.value(),
            "questions": []
        }
        self.accept()

# =============================================================================
#    MAIN APPLICATION
# =============================================================================
class ExamGeneratorApp(QWidget):
    def __init__(self, metadata):
        super().__init__()
        self.metadata = metadata 
        self.setWindowTitle(f"Pro Exam Generator - {metadata['school']}")
        self.resize(1100, 850)
        self.sections = [] 
        self.editing_coords = None 
        
        # --- AUTO LOAD JJN FONT ---
        self.urdu_font_path = "JNN.ttf" if os.path.exists("JNN.ttf") else None
        if not self.urdu_font_path and os.path.exists("jnn.ttf"): 
            self.urdu_font_path = "jnn.ttf"
            
        if not self.urdu_font_path:
            print("Notice: 'JNN.ttf' not found. Urdu will use system fallback.")
        
        # --- MAIN UI SETUP ---
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0,0,0,0)
        
        # Header Bar
        header_bar = QFrame()
        header_bar.setStyleSheet("background-color: #2c3e50; border-bottom: 4px solid #f1c40f;")
        hb_layout = QHBoxLayout(header_bar)
        lbl_title = QLabel(f"DAR-E-ARQAM EXAM CREATOR | {metadata['class']} - {metadata['subject']}")
        lbl_title.setStyleSheet("color: white; font-size: 16px; font-weight: bold;")
        hb_layout.addWidget(lbl_title)
        main_layout.addWidget(header_bar)

        # Scroll Area
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        content_widget = QWidget()
        self.layout = QVBoxLayout(content_widget)
        self.layout.setSpacing(15)
        self.layout.setContentsMargins(20, 20, 20, 20)
        scroll.setWidget(content_widget)
        
        # UI Blocks
        self.create_top_actions()
        self.create_info_display() 
        self.create_section_ui()
        self.create_question_ui()
        
        # Preview Area
        self.preview_list = QListWidget()
        self.preview_list.setStyleSheet("border: 1px solid #bdc3c7; background: #ecf0f1; font-size: 13px;")
        self.preview_list.itemClicked.connect(self.load_question_for_edit)
        
        self.layout.addWidget(QLabel("<b>Structure Preview:</b>"))
        self.layout.addWidget(self.preview_list)
        
        # Export Button
        export_btn = QPushButton(f"EXPORT PDF ({metadata['subject']})")
        export_btn.setCursor(Qt.PointingHandCursor)
        export_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60; color: white; padding: 15px; 
                font-weight: bold; font-size: 16px; border-radius: 5px;
            }
            QPushButton:hover { background-color: #2ecc71; }
        """)
        export_btn.clicked.connect(self.export_pdf)
        self.layout.addWidget(export_btn)
        
        main_layout.addWidget(scroll)
        self.setLayout(main_layout)
        
        # Styles
        self.setStyleSheet("""
            QGroupBox { font-weight: bold; border: 1px solid #bdc3c7; border-radius: 5px; margin-top: 10px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 5px; left: 10px; }
            QLineEdit, QComboBox, QTextEdit, QSpinBox { padding: 6px; border: 1px solid #bdc3c7; border-radius: 4px; background: white; }
            QPushButton { background-color: #34495e; color: white; border-radius: 4px; padding: 8px; }
            QPushButton:hover { background-color: #4b6584; }
        """)

    # --- URDU PROCESSING ---
    def process_text(self, text):
        """Returns: (processed_text, font_property, is_urdu_bool)"""
        if not HAS_URDU_LIB or not text:
            return text, None, False
            
        # Regex to detect Urdu/Arabic chars
        is_urdu = bool(re.search('[\u0600-\u06FF]', text))
        
        if is_urdu:
            # Reshape (connect letters)
            reshaped_text = arabic_reshaper.reshape(text)
            # Reorder (Right-to-Left visual flip)
            bidi_text = get_display(reshaped_text)
            
            if self.urdu_font_path:
                prop = fm.FontProperties(fname=self.urdu_font_path)
            else:
                prop = fm.FontProperties(family=['Arial', 'Segoe UI', 'Tahoma'])
            return bidi_text, prop, True
        else:
            return text, None, False

    # --- UI COMPONENTS ---
    def create_top_actions(self):
        bar = QHBoxLayout()
        save_btn = QPushButton("Save Progress (.csv)")
        save_btn.clicked.connect(self.save_csv)
        save_btn.setStyleSheet("background-color: #e67e22;")
        
        load_btn = QPushButton("Load Progress (.csv)")
        load_btn.clicked.connect(self.load_csv)
        load_btn.setStyleSheet("background-color: #8e44ad;")
        
        bar.addWidget(save_btn)
        bar.addWidget(load_btn)
        bar.addStretch()
        self.layout.addLayout(bar)

    def create_info_display(self):
        frame = QFrame()
        frame.setStyleSheet("background-color: #f8f9fa; border: 1px solid #ddd; border-radius: 5px;")
        l = QHBoxLayout(frame)
        info_str = (f"<b>Class:</b> {self.metadata['class']} &nbsp;|&nbsp; "
                    f"<b>Subject:</b> {self.metadata['subject']} &nbsp;|&nbsp; "
                    f"<b>Time:</b> {self.metadata['time']} &nbsp;|&nbsp; "
                    f"<b>Marks:</b> {self.metadata['marks']}")
        lbl = QLabel(info_str)
        lbl.setStyleSheet("color: #333; font-size: 14px;")
        l.addWidget(lbl)
        self.layout.addWidget(frame)

    def create_section_ui(self):
        group = QGroupBox("1. Manage Sections")
        layout = QHBoxLayout()
        
        lbl = QLabel("Click to add a new section (Title, Marks, Choices)")
        
        add_btn = QPushButton("+ Add New Section")
        add_btn.setStyleSheet("background-color: #27ae60; font-weight: bold; padding: 10px;")
        add_btn.clicked.connect(self.open_section_dialog)
        
        layout.addWidget(lbl)
        layout.addStretch()
        layout.addWidget(add_btn)
        group.setLayout(layout)
        self.layout.addWidget(group)

    def create_question_ui(self):
        group = QGroupBox("2. Add Questions")
        layout = QVBoxLayout()
        
        h = QHBoxLayout()
        self.section_selector = QComboBox()
        self.q_type = QComboBox()
        self.q_type.addItems(["MCQ", "Short/Long Question", "Match Columns"])
        self.q_type.currentTextChanged.connect(self.toggle_inputs)
        
        h.addWidget(QLabel("Add to Section:"))
        h.addWidget(self.section_selector, 1)
        h.addWidget(QLabel("Type:"))
        h.addWidget(self.q_type, 1)
        layout.addLayout(h)
        
        self.q_text = QTextEdit()
        self.q_text.setMaximumHeight(80)
        self.q_text.setPlaceholderText("Enter Question Text (Urdu or English)...")
        layout.addWidget(self.q_text)
        
        # MCQ
        self.mcq_widget = QWidget()
        l_mcq = QHBoxLayout(self.mcq_widget)
        l_mcq.setContentsMargins(0,0,0,0)
        self.opt_inputs = []
        for i in range(4):
            le = QLineEdit()
            le.setPlaceholderText(f"Option {chr(65+i)}")
            self.opt_inputs.append(le)
            l_mcq.addWidget(le)
        layout.addWidget(self.mcq_widget)
        
        # Match Columns
        self.match_widget = QWidget()
        l_match = QHBoxLayout(self.match_widget)
        self.col_a = QTextEdit()
        self.col_a.setPlaceholderText("Column A (one per line)")
        self.col_a.setMaximumHeight(80)
        self.col_b = QTextEdit()
        self.col_b.setPlaceholderText("Column B (one per line)")
        self.col_b.setMaximumHeight(80)
        l_match.addWidget(self.col_a)
        l_match.addWidget(self.col_b)
        layout.addWidget(self.match_widget)
        self.match_widget.hide()
        
        # Buttons
        btn_row = QHBoxLayout()
        self.add_btn = QPushButton("Save Question")
        self.add_btn.clicked.connect(self.save_question)
        self.add_btn.setStyleSheet("background-color: #2980b9; font-weight: bold;")
        self.cancel_btn = QPushButton("Cancel Edit")
        self.cancel_btn.hide()
        self.cancel_btn.clicked.connect(self.reset_form)
        self.cancel_btn.setStyleSheet("background-color: #e74c3c;")
        btn_row.addWidget(self.add_btn)
        btn_row.addWidget(self.cancel_btn)
        layout.addLayout(btn_row)
        
        group.setLayout(layout)
        self.layout.addWidget(group)

    # --- LOGIC ---
    def toggle_inputs(self, txt):
        self.mcq_widget.setVisible(txt == "MCQ")
        self.match_widget.setVisible(txt == "Match Columns")

    def open_section_dialog(self):
        """Opens the pop-up to add a section with choice features."""
        dlg = SectionSetupDialog(self)
        if dlg.exec():
            self.sections.append(dlg.section_data)
            self.refresh_ui()

    def refresh_ui(self):
        curr = self.section_selector.currentIndex()
        self.section_selector.clear()
        self.preview_list.clear()
        
        for s in self.sections:
            self.section_selector.addItem(f"{s['name']}")
            
            # Format: (Marks x Count = Total)
            mark_str = f"({s['marks_per_q']} x {s['attempt_count']} = {s['total_marks']})"
            header_str = f"--- {s['name']} : {s['desc']} {mark_str} ---"
            
            self.preview_list.addItem(header_str)
            for i, q in enumerate(s['questions']):
                self.preview_list.addItem(f"    {i+1}. {q['text'][:40]}...")
                
        if curr >= 0 and curr < self.section_selector.count():
            self.section_selector.setCurrentIndex(curr)

    def save_question(self):
        # 1. Check if section is selected
        idx = self.section_selector.currentIndex()
        if idx < 0: 
            QMessageBox.warning(self, "Warning", "Please create and select a section first.")
            return

        # 2. VALIDATION: Check if question text is empty
        raw_text = self.q_text.toPlainText().strip()
        if not raw_text:
            QMessageBox.warning(self, "Input Error", "Question text cannot be empty!")
            return

        qtype = self.q_type.currentText()
        q = {"type": qtype, "text": raw_text}
        
        if qtype == "MCQ": 
            q["options"] = [o.text() for o in self.opt_inputs]
        elif qtype == "Match Columns":
            q["col_a"] = [x.strip() for x in self.col_a.toPlainText().split('\n') if x.strip()]
            q["col_b"] = [x.strip() for x in self.col_b.toPlainText().split('\n') if x.strip()]
        
        if self.editing_coords:
            si, qi = self.editing_coords
            self.sections[si]['questions'][qi] = q
        else:
            self.sections[idx]['questions'].append(q)
        
        self.reset_form()
        self.refresh_ui()

    def load_question_for_edit(self, item):
        row = self.preview_list.row(item)
        curr = 0
        for s_i, sec in enumerate(self.sections):
            curr += 1 
            if row == curr - 1: return # Clicked header
            for q_i, q in enumerate(sec['questions']):
                if row == curr: self.start_edit(s_i, q_i); return
                curr += 1

    def start_edit(self, s_i, q_i):
        self.editing_coords = (s_i, q_i)
        q = self.sections[s_i]['questions'][q_i]
        self.section_selector.setCurrentIndex(s_i)
        self.q_type.setCurrentText(q['type'])
        self.q_text.setText(q['text'])
        if q['type'] == "MCQ":
            for i, o in enumerate(q.get('options', [])):
                if i < 4: self.opt_inputs[i].setText(o)
        elif q['type'] == "Match Columns":
            self.col_a.setText("\n".join(q.get('col_a', [])))
            self.col_b.setText("\n".join(q.get('col_b', [])))
        self.add_btn.setText("Update"); self.cancel_btn.show()

    def reset_form(self):
        self.editing_coords = None
        self.q_text.clear(); self.col_a.clear(); self.col_b.clear()
        for o in self.opt_inputs: o.clear()
        self.add_btn.setText("Save Question"); self.cancel_btn.hide()

    def save_csv(self):
        fn, _ = QFileDialog.getSaveFileName(self, "Save", "", "CSV (*.csv)")
        if fn:
            try:
                with open(fn, 'w', newline='', encoding='utf-8') as f:
                    w = csv.writer(f)
                    md = self.metadata
                    w.writerow(["META", md['school'], md['test'], md['class'], md['subject'], md['time'], md['marks']])
                    for s in self.sections:
                        # Updated SEC row to include marks per q and attempt count
                        w.writerow(["SEC", s['name'], s['desc'], s['marks_per_q'], s['attempt_count']])
                        for q in s['questions']:
                            row = ["Q", q['type'], q['text']]
                            if q['type'] == "MCQ": row += q['options']
                            elif q['type'] == "Match Columns": row += ["|".join(q['col_a']), "|".join(q['col_b'])]
                            w.writerow(row)
                QMessageBox.information(self, "Saved", "Done.")
            except Exception as e: 
                QMessageBox.critical(self, "Error", str(e))

    def load_csv(self):
        fn, _ = QFileDialog.getOpenFileName(self, "Load", "", "CSV (*.csv)")
        if fn:
            self.sections = []
            curr_sec = None
            try:
                with open(fn, 'r', encoding='utf-8') as f:
                    r = csv.reader(f)
                    for row in r:
                        if not row: continue
                        if row[0] == "META":
                            self.metadata = {
                                "school": row[1], "test": row[2], "class": row[3],
                                "subject": row[4], "time": row[5], "marks": row[6]
                            }
                        elif row[0] == "SEC":
                            # Handle Backward Compatibility or New Format
                            marks_pq = int(row[3])
                            att_cnt = int(row[4]) if len(row) > 4 else 1 # Default if old file
                            
                            curr_sec = {
                                "name": row[1], 
                                "desc": row[2], 
                                "marks_per_q": marks_pq,
                                "attempt_count": att_cnt,
                                "total_marks": marks_pq * att_cnt,
                                "questions": []
                            }
                            self.sections.append(curr_sec)
                        elif row[0] == "Q" and curr_sec:
                            q = {"type": row[1], "text": row[2]}
                            if q['type'] == "MCQ": q['options'] = row[3:]
                            elif q['type'] == "Match Columns": q['col_a'] = row[3].split('|'); q['col_b'] = row[4].split('|')
                            curr_sec['questions'].append(q)
                self.refresh_ui()
            except Exception as e: 
                print(e)
                QMessageBox.warning(self, "Load Error", "Could not load file completely.")

    # =========================================================================
    #    PDF EXPORT LOGIC
    # =========================================================================
    def export_pdf(self):
        fn, _ = QFileDialog.getSaveFileName(self, "Export PDF", f"{self.metadata['subject']}_Exam.pdf", "PDF (*.pdf)")
        if not fn: return

        try:
            # === PAGE CONSTANTS ===
            PAGE_W, PAGE_H = 8.27, 11.69
            MARGIN_X, MARGIN_TOP, MARGIN_BTM = 0.5, 0.5, 0.5
            CONTENT_W = PAGE_W - (2 * MARGIN_X)
            
            FS_HEADER, FS_SUB, FS_BODY = 16, 12, 11
            LH = 0.25 # Line Height

            with PdfPages(fn) as pdf:
                def new_page():
                    fig = plt.figure(figsize=(PAGE_W, PAGE_H))
                    ax = fig.add_axes([0, 0, 1, 1]) 
                    ax.set_xlim(0, PAGE_W); ax.set_ylim(0, PAGE_H); ax.axis('off')
                    return fig, ax

                def draw_text(ax, x, y, txt, fs=10, weight='normal', align='left', v_align='baseline', force_rtl_check=True):
                    final_txt, font_prop, is_urdu = self.process_text(txt)
                    
                    eff_align = align
                    eff_x = x
                    eff_weight = weight
                    eff_fs = fs

                    if is_urdu and force_rtl_check:
                        eff_weight = 'normal' 
                        eff_fs += 2 
                        
                        if align == 'left':
                            eff_align = 'right'
                            eff_x = PAGE_W - x 
                        elif align == 'right':
                            eff_align = 'left'
                            eff_x = PAGE_W - x
                    
                    elif is_urdu and not force_rtl_check:
                        eff_weight = 'normal'
                        eff_fs += 2
                    
                    kwargs = {'fontsize': eff_fs, 'fontweight': eff_weight, 'ha': eff_align, 'va': v_align}
                    if font_prop: kwargs['fontproperties'] = font_prop
                    
                    ax.text(eff_x, y, final_txt, **kwargs)
                    return is_urdu

                fig, ax = new_page()
                cursor_y = PAGE_H - MARGIN_TOP

                # --- HEADER BOX ---
                H_BOX_H = 2.0
                p_box = FancyBboxPatch((MARGIN_X, cursor_y - H_BOX_H), CONTENT_W, H_BOX_H,
                                     boxstyle="round,pad=0.1", ec="black", fc="white", lw=2)
                ax.add_patch(p_box)
                
                center_x = PAGE_W / 2
                draw_text(ax, center_x, cursor_y - 0.3, self.metadata['school'], FS_HEADER, 'bold', 'center')
                draw_text(ax, center_x, cursor_y - 0.6, self.metadata['test'], FS_SUB, 'bold', 'center')
                
                rule_y = cursor_y - 0.8
                ax.plot([MARGIN_X + 0.2, PAGE_W - MARGIN_X - 0.2], [rule_y, rule_y], color='black', lw=1)
                
                meta_y = rule_y - 0.3
                draw_text(ax, MARGIN_X + 0.2, meta_y, f"Class: {self.metadata['class']}", FS_BODY)
                draw_text(ax, center_x + 0.5, meta_y, f"Time: {self.metadata['time']}", FS_BODY)
                
                meta_y -= 0.25
                draw_text(ax, MARGIN_X + 0.2, meta_y, f"Subject: {self.metadata['subject']}", FS_BODY)
                draw_text(ax, center_x + 0.5, meta_y, f"Marks: {self.metadata['marks']}", FS_BODY)

                # --- STUDENT NAME ROW ---
                name_y = meta_y - 0.35
                draw_text(ax, MARGIN_X + 0.2, name_y, "Student Name: ____________________________", FS_BODY)
                draw_text(ax, center_x + 0.5, name_y, "Roll No: ____________", FS_BODY)
                cursor_y -= (H_BOX_H + 0.3)

                # --- SECTIONS ---
                for sec in self.sections:
                    if cursor_y < MARGIN_BTM + 1.0:
                        pdf.savefig(fig); plt.close(); fig, ax = new_page(); cursor_y = PAGE_H - MARGIN_TOP
                    
                    # Section Header Bar
                    SEC_BOX_H = 0.4
                    p_sec = FancyBboxPatch((MARGIN_X, cursor_y - SEC_BOX_H), CONTENT_W, SEC_BOX_H,
                                           boxstyle="round,pad=0.05", ec="black", fc="#bdc3c7", lw=1)
                    ax.add_patch(p_sec)
                    
                    sec_text_y = cursor_y - 0.25
                    title = f"{sec['name']}   {sec['desc']}"
                    
                    # --- NEW FORMAT: (Marks x Count = Total) ---
                    marks_txt = f"({sec['marks_per_q']} x {sec['attempt_count']} = {sec['total_marks']})"
                    
                    _, _, is_sec_urdu = self.process_text(sec['desc'])
                    
                    if is_sec_urdu:
                          draw_text(ax, PAGE_W - MARGIN_X - 0.1, sec_text_y, title, FS_BODY, 'normal', 'right', 'center', False)
                          draw_text(ax, MARGIN_X + 0.1, sec_text_y, marks_txt, FS_BODY, 'bold', 'left', 'center', False)
                    else:
                          draw_text(ax, MARGIN_X + 0.1, sec_text_y, title, FS_BODY, 'bold', 'left', 'center', False)
                          draw_text(ax, PAGE_W - MARGIN_X - 0.1, sec_text_y, marks_txt, FS_BODY, 'bold', 'right', 'center', False)

                    cursor_y -= (SEC_BOX_H + 0.25)

                    # --- QUESTIONS ---
                    for idx, q in enumerate(sec['questions']):
                        q_num = f"{idx+1}."
                        lines = textwrap.wrap(q['text'], width=85)
                        req_h = (len(lines) * LH) + 0.2
                        if q['type'] == "MCQ": req_h += 0.8
                        if q['type'] == "Match Columns": req_h += 1.5
                        
                        if cursor_y - req_h < MARGIN_BTM:
                            pdf.savefig(fig); plt.close(); fig, ax = new_page(); cursor_y = PAGE_H - MARGIN_TOP
                        
                        _, _, is_q_urdu = self.process_text(q['text'])
                        
                        if is_q_urdu:
                            draw_text(ax, MARGIN_X, cursor_y, q_num, FS_BODY, 'bold', 'left', 'top', force_rtl_check=False)
                            text_x_anchor = PAGE_W - MARGIN_X - 0.2 
                            
                            for line in lines:
                                draw_text(ax, text_x_anchor, cursor_y, line, FS_BODY, 'normal', 'right', 'top', force_rtl_check=False)
                                cursor_y -= LH
                                
                            text_x_anchor_for_opts = text_x_anchor 
                        else:
                            draw_text(ax, MARGIN_X, cursor_y, q_num, FS_BODY, 'bold', 'left', 'top', False)
                            text_x_anchor = MARGIN_X + 0.6
                            for line in lines:
                                draw_text(ax, text_x_anchor, cursor_y, line, FS_BODY, 'normal', 'left', 'top', False)
                                cursor_y -= LH
                            
                            text_x_anchor_for_opts = text_x_anchor

                        cursor_y -= 0.1
                        
                        # Options
                        if q['type'] == "MCQ":
                            opts = q.get('options', [])
                            opt_y = cursor_y
                            
                            if is_q_urdu:
                                if len(opts) > 0: draw_text(ax, text_x_anchor_for_opts, opt_y, f"{opts[0]} (a)", FS_BODY, 'normal', 'right', 'baseline', False)
                                if len(opts) > 2: draw_text(ax, text_x_anchor_for_opts, opt_y - LH, f"{opts[2]} (c)", FS_BODY, 'normal', 'right', 'baseline', False)
                                
                                if len(opts) > 1: draw_text(ax, text_x_anchor_for_opts - 3.5, opt_y, f"{opts[1]} (b)", FS_BODY, 'normal', 'right', 'baseline', False)
                                if len(opts) > 3: draw_text(ax, text_x_anchor_for_opts - 3.5, opt_y - LH, f"{opts[3]} (d)", FS_BODY, 'normal', 'right', 'baseline', False)
                                
                                cursor_y = opt_y - (2 * LH) - 0.1

                            else:
                                if len(opts) > 0: draw_text(ax, text_x_anchor_for_opts, opt_y, f"(a) {opts[0]}", FS_BODY, 'normal', 'left', 'baseline', False)
                                if len(opts) > 1: draw_text(ax, text_x_anchor_for_opts + 3.5, opt_y, f"(b) {opts[1]}", FS_BODY, 'normal', 'left', 'baseline', False)
                                opt_y -= LH
                                if len(opts) > 2: draw_text(ax, text_x_anchor_for_opts, opt_y, f"(c) {opts[2]}", FS_BODY, 'normal', 'left', 'baseline', False)
                                if len(opts) > 3: draw_text(ax, text_x_anchor_for_opts + 3.5, opt_y, f"(d) {opts[3]}", FS_BODY, 'normal', 'left', 'baseline', False)
                                cursor_y = opt_y - 0.3
                            
                        elif q['type'] == "Match Columns":
                            col_a = q.get('col_a', [])
                            col_b = q.get('col_b', [])
                            
                            draw_text(ax, MARGIN_X + 1.0, cursor_y, "Column A", FS_BODY, 'bold')
                            draw_text(ax, PAGE_W - MARGIN_X - 1.0, cursor_y, "Column B", FS_BODY, 'bold', 'right')
                            cursor_y -= LH
                            
                            ax.plot([MARGIN_X, PAGE_W - MARGIN_X], [cursor_y + 0.1, cursor_y + 0.1], color='black', lw=0.5)
                            
                            rows = max(len(col_a), len(col_b))
                            for r in range(rows):
                                ta = col_a[r] if r < len(col_a) else ""
                                tb = col_b[r] if r < len(col_b) else ""
                                draw_text(ax, MARGIN_X + 1.0, cursor_y, ta, FS_BODY)
                                draw_text(ax, PAGE_W - MARGIN_X - 1.0, cursor_y, tb, FS_BODY, 'normal', 'right')
                                cursor_y -= LH
                            cursor_y -= 0.2
                        else:
                            cursor_y -= 0.15

                draw_text(ax, PAGE_W/2, MARGIN_BTM, "End of Question Paper", 8, 'normal', 'center')
                pdf.savefig(fig)
                plt.close()

            QMessageBox.information(self, "Success", f"PDF Exported!\nSaved to: {fn}")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))
            print(e)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    setup = SetupDetailsDialog()
    if setup.exec() == QDialog.Accepted:
        window = ExamGeneratorApp(setup.data)
        window.show()
        sys.exit(app.exec())
