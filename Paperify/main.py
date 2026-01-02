import sys
import csv
import textwrap
import re
import os
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QPushButton, QComboBox,
    QTreeWidget, QTreeWidgetItem, QMessageBox, QLineEdit, QSpinBox, 
    QFormLayout, QGroupBox, QScrollArea, QFileDialog,
    QDialog, QDialogButtonBox, QFrame, QMenu, QAbstractItemView, QSplitter
)
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QColor, QAction, QFont, QIcon

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
    print("Warning: 'arabic-reshaper' or 'python-bidi' not installed.")

# --- PDF SETTINGS ---
plt.rcParams['font.family'] = 'serif'
plt.rcParams['mathtext.fontset'] = 'cm' 
plt.rcParams['axes.unicode_minus'] = False 

# =============================================================================
#    DIALOG: STARTUP DETAILS
# =============================================================================
class SetupDetailsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Exam Configuration")
        self.setModal(True)
        self.resize(500, 400)
        self.data = {}
        
        # Styles
        self.setStyleSheet("""
            QDialog { background-color: #f4f6f9; }
            QLabel { font-size: 14px; color: #34495e; }
            QLineEdit { padding: 8px; border: 1px solid #bdc3c7; border-radius: 4px; }
            QLineEdit:focus { border: 2px solid #3498db; }
            QPushButton { padding: 8px 15px; border-radius: 4px; font-weight: bold; }
        """)
        
        layout = QVBoxLayout()
        
        # Header
        header = QFrame()
        header.setStyleSheet("background-color: #2c3e50; border-radius: 5px;")
        hl = QVBoxLayout(header)
        lbl_brand = QLabel("DAR-E-ARQAM SCHOOL")
        lbl_brand.setAlignment(Qt.AlignCenter)
        lbl_brand.setStyleSheet("font-size: 22px; font-weight: bold; color: white;")
        hl.addWidget(lbl_brand)
        layout.addWidget(header)
        
        form_layout = QFormLayout()
        form_layout.setSpacing(15)
        
        self.inp_school = QLineEdit("DAR-E-ARQAM SCHOOL")
        self.inp_test = QLineEdit("Monthly Test")
        self.inp_class = QLineEdit()
        self.inp_class.setPlaceholderText("e.g. 9th")
        self.inp_subject = QLineEdit()
        self.inp_subject.setPlaceholderText("e.g. Physics")
        self.inp_time = QLineEdit("1 Hr 30 Mins")
        self.inp_marks = QLineEdit("50")
        
        form_layout.addRow("School Name:", self.inp_school)
        form_layout.addRow("Test Title:", self.inp_test)
        form_layout.addRow("Class:", self.inp_class)
        form_layout.addRow("Subject:", self.inp_subject)
        form_layout.addRow("Total Time:", self.inp_time)
        form_layout.addRow("Total Marks:", self.inp_marks)
        
        layout.addLayout(form_layout)
        
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.button(QDialogButtonBox.Ok).setStyleSheet("background-color: #27ae60; color: white;")
        buttons.button(QDialogButtonBox.Cancel).setStyleSheet("background-color: #e74c3c; color: white;")
        
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
#    DIALOG: SECTION SETUP (Edit/Add)
# =============================================================================
class SectionSetupDialog(QDialog):
    def __init__(self, parent=None, existing_data=None):
        super().__init__(parent)
        self.setWindowTitle("Section Configuration")
        self.setModal(True)
        self.resize(450, 350)
        self.section_data = None
        
        self.setStyleSheet("""
            QLabel { font-size: 13px; font-weight: bold; }
            QLineEdit, QSpinBox { padding: 6px; border: 1px solid #bdc3c7; border-radius: 4px; }
        """)

        layout = QVBoxLayout()
        form = QFormLayout()
        
        self.inp_title = QLineEdit()
        self.inp_title.setPlaceholderText("e.g. Section B")
        
        self.inp_desc = QLineEdit()
        self.inp_desc.setPlaceholderText("e.g. Attempt any 5 questions")

        self.spin_marks_per_q = QSpinBox()
        self.spin_marks_per_q.setRange(1, 100)
        self.spin_marks_per_q.setValue(2)

        self.spin_attempt_count = QSpinBox()
        self.spin_attempt_count.setRange(1, 100)
        self.spin_attempt_count.setValue(5)

        self.lbl_total_calc = QLabel("Total: 10 Marks")
        self.lbl_total_calc.setStyleSheet("color: #27ae60; font-size: 14px;")

        # Events
        self.spin_marks_per_q.valueChanged.connect(self.update_total)
        self.spin_attempt_count.valueChanged.connect(self.update_total)

        form.addRow("Section Title:", self.inp_title)
        form.addRow("Instruction:", self.inp_desc)
        form.addRow("Marks per Question:", self.spin_marks_per_q)
        form.addRow("Questions to Attempt:", self.spin_attempt_count)
        form.addRow("Calculated Total:", self.lbl_total_calc)

        layout.addLayout(form)
        
        # Pre-fill if editing
        if existing_data:
            self.inp_title.setText(existing_data.get('name', ''))
            self.inp_desc.setText(existing_data.get('desc', ''))
            self.spin_marks_per_q.setValue(int(existing_data.get('marks_per_q', 2)))
            self.spin_attempt_count.setValue(int(existing_data.get('attempt_count', 5)))
            self.section_questions = existing_data.get('questions', []) # Preserve questions
        else:
            self.section_questions = []

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
            "questions": self.section_questions
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
        self.resize(1200, 900)
        
        # Data storage
        self.sections = [] 
        self.editing_q_ptr = None # Pointer to (section_index, question_index)
        
        # Font Loading
        self.urdu_font_path = "JNN.ttf" if os.path.exists("JNN.ttf") else None
        if not self.urdu_font_path and os.path.exists("jnn.ttf"): self.urdu_font_path = "jnn.ttf"
        
        self.apply_styles()
        self.init_ui()

    def apply_styles(self):
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', sans-serif; font-size: 14px; color: #2c3e50; }
            
            /* Group Boxes */
            QGroupBox { 
                font-weight: bold; border: 2px solid #bdc3c7; border-radius: 6px; 
                margin-top: 20px; padding-top: 15px; background-color: white;
            }
            QGroupBox::title { 
                subcontrol-origin: margin; subcontrol-position: top left; 
                padding: 0 5px; left: 10px; color: #2980b9; 
            }

            /* Inputs */
            QLineEdit, QComboBox, QTextEdit, QSpinBox { 
                padding: 8px; border: 1px solid #bdc3c7; border-radius: 4px; background: #fdfdfd; 
            }
            QLineEdit:focus, QTextEdit:focus { border: 1px solid #3498db; background: white; }

            /* Buttons */
            QPushButton { 
                background-color: #34495e; color: white; border-radius: 4px; padding: 8px 16px; 
            }
            QPushButton:hover { background-color: #2c3e50; border: 1px solid #f1c40f; }
            
            /* Tree Widget */
            QTreeWidget { border: 1px solid #bdc3c7; background-color: #ecf0f1; }
            QTreeWidget::item { padding: 5px; }
            QTreeWidget::item:selected { background-color: #3498db; color: white; }
        """)

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(0,0,0,0)
        
        # --- HEADER ---
        header = QFrame()
        header.setStyleSheet("background-color: #2c3e50; border-bottom: 4px solid #f1c40f;")
        hl = QHBoxLayout(header)
        title = QLabel(f"EXAM CREATOR | {self.metadata['class']} - {self.metadata['subject']}")
        title.setStyleSheet("color: white; font-size: 18px; font-weight: bold;")
        hl.addWidget(title)
        
        # Export Button in Header
        btn_export = QPushButton(f"EXPORT PDF")
        btn_export.setCursor(Qt.PointingHandCursor)
        btn_export.setStyleSheet("background-color: #e74c3c; font-weight: bold; border: none;")
        btn_export.clicked.connect(self.export_pdf)
        hl.addStretch()
        hl.addWidget(btn_export)
        main_layout.addWidget(header)

        # --- BODY CONTENT (Splitter) ---
        splitter = QSplitter(Qt.Horizontal)
        
        # LEFT PANE: Inputs
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        left_widget = QWidget()
        self.left_layout = QVBoxLayout(left_widget)
        self.left_layout.setSpacing(15)
        self.left_layout.setContentsMargins(15, 15, 15, 15)
        
        self.create_top_actions()
        self.create_info_panel()
        self.create_question_input_ui()
        
        self.left_layout.addStretch()
        scroll_area.setWidget(left_widget)
        
        # RIGHT PANE: Structure Tree
        right_widget = QWidget()
        rl = QVBoxLayout(right_widget)
        rl.setContentsMargins(10, 15, 10, 10)
        
        rl.addWidget(QLabel("<b>Exam Structure (Drag to Reorder)</b>"))
        
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDragDropMode(QAbstractItemView.InternalMove)
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.open_context_menu)
        self.tree.itemDoubleClicked.connect(self.on_tree_double_click)
        
        # Add Section Button
        btn_add_sec = QPushButton("+ Add New Section")
        btn_add_sec.setStyleSheet("background-color: #27ae60; width: 100%;")
        btn_add_sec.clicked.connect(lambda: self.open_section_dialog(None))
        
        rl.addWidget(self.tree)
        rl.addWidget(btn_add_sec)
        
        splitter.addWidget(scroll_area)
        splitter.addWidget(right_widget)
        splitter.setSizes([600, 500])
        
        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

    # --- UI COMPONENTS ---
    def create_top_actions(self):
        bar = QHBoxLayout()
        save_btn = QPushButton("Save Progress (.csv)")
        save_btn.clicked.connect(self.save_csv)
        save_btn.setStyleSheet("background-color: #d35400;")
        
        load_btn = QPushButton("Load Progress (.csv)")
        load_btn.clicked.connect(self.load_csv)
        load_btn.setStyleSheet("background-color: #8e44ad;")
        
        bar.addWidget(save_btn)
        bar.addWidget(load_btn)
        bar.addStretch()
        self.left_layout.addLayout(bar)

    def create_info_panel(self):
        grp = QGroupBox("Exam Metadata")
        l = QHBoxLayout()
        info = (f"<b>Class:</b> {self.metadata['class']} | "
                f"<b>Time:</b> {self.metadata['time']} | "
                f"<b>Marks:</b> {self.metadata['marks']}")
        lbl = QLabel(info)
        l.addWidget(lbl)
        grp.setLayout(l)
        self.left_layout.addWidget(grp)

    def create_question_input_ui(self):
        grp = QGroupBox("Question Editor")
        layout = QVBoxLayout()
        
        # Inputs
        h = QHBoxLayout()
        h.addWidget(QLabel("Section:"))
        self.cb_sections = QComboBox() # To select which section to add to
        h.addWidget(self.cb_sections, 1)
        h.addWidget(QLabel("Type:"))
        self.q_type = QComboBox()
        self.q_type.addItems(["MCQ", "Short/Long Question", "Match Columns"])
        self.q_type.currentTextChanged.connect(self.toggle_inputs)
        h.addWidget(self.q_type, 1)
        layout.addLayout(h)
        
        self.q_text = QTextEdit()
        self.q_text.setMaximumHeight(80)
        self.q_text.setPlaceholderText("Enter Question Text (Urdu/English)...")
        layout.addWidget(self.q_text)
        
        # MCQ Options
        self.mcq_widget = QWidget()
        l_mcq = QVBoxLayout(self.mcq_widget)
        l_mcq.setContentsMargins(0,0,0,0)
        l_mcq.addWidget(QLabel("Options:"))
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
        self.col_a.setPlaceholderText("Column A (Lines)")
        self.col_a.setMaximumHeight(80)
        self.col_b = QTextEdit()
        self.col_b.setPlaceholderText("Column B (Lines)")
        self.col_b.setMaximumHeight(80)
        l_match.addWidget(self.col_a)
        l_match.addWidget(self.col_b)
        layout.addWidget(self.match_widget)
        self.match_widget.hide()
        
        # Buttons
        btn_row = QHBoxLayout()
        self.btn_save_q = QPushButton("Add Question")
        self.btn_save_q.clicked.connect(self.save_question_input)
        self.btn_save_q.setStyleSheet("background-color: #2980b9; font-size: 15px;")
        
        self.btn_cancel = QPushButton("Cancel Edit")
        self.btn_cancel.hide()
        self.btn_cancel.clicked.connect(self.reset_editor)
        self.btn_cancel.setStyleSheet("background-color: #7f8c8d;")
        
        btn_row.addWidget(self.btn_save_q)
        btn_row.addWidget(self.btn_cancel)
        layout.addLayout(btn_row)
        
        grp.setLayout(layout)
        self.left_layout.addWidget(grp)

    def toggle_inputs(self, txt):
        self.mcq_widget.setVisible(txt == "MCQ")
        self.match_widget.setVisible(txt == "Match Columns")

    # --- LOGIC & TREE MANAGEMENT ---
    
    def refresh_sections_combo(self):
        curr = self.cb_sections.currentText()
        self.cb_sections.clear()
        for s in self.sections:
            self.cb_sections.addItem(s['name'])
        self.cb_sections.setCurrentText(curr)

    def rebuild_tree(self):
        """Syncs the visual TreeWidget with self.sections data"""
        self.tree.clear()
        for s_idx, sec in enumerate(self.sections):
            sec_item = QTreeWidgetItem(self.tree)
            sec_item.setText(0, f"{sec['name']} - {sec['desc']} ({sec['total_marks']} marks)")
            sec_item.setData(0, Qt.UserRole, "SECTION")
            sec_item.setData(0, Qt.UserRole+1, s_idx)
            
            font = QFont(); font.setBold(True); sec_item.setFont(0, font)
            sec_item.setBackground(0, QColor("#dfe6e9"))
            
            for q_idx, q in enumerate(sec['questions']):
                q_item = QTreeWidgetItem(sec_item)
                display_text = q['text'][:50] + "..." if len(q['text']) > 50 else q['text']
                q_item.setText(0, f"Q{q_idx+1}: {display_text}")
                q_item.setData(0, Qt.UserRole, "QUESTION")
                q_item.setData(0, Qt.UserRole+1, q) # Store actual question data temporarily
            
            sec_item.setExpanded(True)
        self.refresh_sections_combo()

    def sync_tree_to_model(self):
        """Scrapes the TreeWidget to update self.sections (handles reordering)"""
        new_sections = []
        root = self.tree.invisibleRootItem()
        
        for i in range(root.childCount()):
            sec_item = root.child(i)
            role = sec_item.data(0, Qt.UserRole)
            
            # If a question was dragged to root, ignore or handle (here we ignore/warn in logic)
            if role != "SECTION": continue
            
            s_idx = sec_item.data(0, Qt.UserRole+1)
            # Retrieve original section metadata (metadata might be stale if reordered, 
            # so ideally we look up by ID or keep existing object. 
            # Simplified: we use the stored index to fetch original metadata)
            
            # However, if we deleted/added, indices change. 
            # Better approach: The tree represents the truth for order.
            # We need to match the tree item back to the original section data.
            original_sec_data = self.sections[s_idx]
            
            new_questions = []
            for j in range(sec_item.childCount()):
                q_item = sec_item.child(j)
                # The question data is stored in UserRole+1
                q_data = q_item.data(0, Qt.UserRole+1)
                new_questions.append(q_data)
            
            original_sec_data['questions'] = new_questions
            new_sections.append(original_sec_data)
            
            # Update index for next time
            sec_item.setData(0, Qt.UserRole+1, i) 
            
        self.sections = new_sections
        self.refresh_sections_combo()

    def open_section_dialog(self, existing_data=None):
        dlg = SectionSetupDialog(self, existing_data)
        if dlg.exec():
            if existing_data:
                # Update logic
                # Find which section we are editing based on object reference or selection
                # Since we pass dict reference, updating it directly works if list structure holds
                existing_data.update(dlg.section_data)
            else:
                self.sections.append(dlg.section_data)
            self.rebuild_tree()

    def save_question_input(self):
        # 1. Sync tree first to ensure indices are current
        self.sync_tree_to_model()
        
        raw_text = self.q_text.toPlainText().strip()
        if not raw_text: return

        q = {
            "type": self.q_type.currentText(),
            "text": raw_text
        }
        
        if q["type"] == "MCQ":
            q["options"] = [o.text() for o in self.opt_inputs]
        elif q["type"] == "Match Columns":
            q["col_a"] = [x for x in self.col_a.toPlainText().split('\n') if x.strip()]
            q["col_b"] = [x for x in self.col_b.toPlainText().split('\n') if x.strip()]

        if self.editing_q_ptr:
            # Update existing
            s_idx, q_idx = self.editing_q_ptr
            self.sections[s_idx]['questions'][q_idx] = q
            self.editing_q_ptr = None
            self.btn_save_q.setText("Add Question")
            self.btn_cancel.hide()
        else:
            # Add new
            sec_name = self.cb_sections.currentText()
            # Find section index
            target = next((s for s in self.sections if s['name'] == sec_name), None)
            if target:
                target['questions'].append(q)
            else:
                QMessageBox.warning(self, "Error", "No Section Selected/Exists")
                return

        self.q_text.clear()
        self.col_a.clear(); self.col_b.clear()
        for o in self.opt_inputs: o.clear()
        self.rebuild_tree()

    def reset_editor(self):
        self.editing_q_ptr = None
        self.btn_save_q.setText("Add Question")
        self.btn_cancel.hide()
        self.q_text.clear()

    # --- TREE INTERACTION ---
    def open_context_menu(self, position):
        item = self.tree.itemAt(position)
        if not item: return
        
        menu = QMenu()
        role = item.data(0, Qt.UserRole)
        
        if role == "SECTION":
            act_edit = QAction("Edit Section Details", self)
            act_edit.triggered.connect(lambda: self.edit_section_from_tree(item))
            menu.addAction(act_edit)
            
            act_del = QAction("Delete Section", self)
            act_del.triggered.connect(lambda: self.delete_item(item))
            menu.addAction(act_del)
        elif role == "QUESTION":
            act_edit = QAction("Edit Question", self)
            act_edit.triggered.connect(lambda: self.load_question_for_edit(item))
            menu.addAction(act_edit)
            
            act_del = QAction("Delete Question", self)
            act_del.triggered.connect(lambda: self.delete_item(item))
            menu.addAction(act_del)
            
        menu.exec(self.tree.viewport().mapToGlobal(position))

    def edit_section_from_tree(self, item):
        s_idx = item.data(0, Qt.UserRole+1)
        self.open_section_dialog(self.sections[s_idx])

    def on_tree_double_click(self, item, col):
        if item.data(0, Qt.UserRole) == "QUESTION":
            self.load_question_for_edit(item)

    def load_question_for_edit(self, item):
        self.sync_tree_to_model() # Save current state
        
        parent = item.parent()
        s_idx = parent.data(0, Qt.UserRole+1)
        q_idx = parent.indexOfChild(item)
        
        q = self.sections[s_idx]['questions'][q_idx]
        
        self.editing_q_ptr = (s_idx, q_idx)
        self.cb_sections.setCurrentIndex(s_idx)
        self.q_type.setCurrentText(q['type'])
        self.q_text.setText(q['text'])
        
        if q['type'] == "MCQ":
            opts = q.get('options', [])
            for i, le in enumerate(self.opt_inputs):
                if i < len(opts): le.setText(opts[i])
        elif q['type'] == "Match Columns":
            self.col_a.setText("\n".join(q.get('col_a', [])))
            self.col_b.setText("\n".join(q.get('col_b', [])))

        self.btn_save_q.setText("Update Question")
        self.btn_cancel.show()

    def delete_item(self, item):
        parent = item.parent()
        if parent: 
            # It's a question
            parent.removeChild(item)
        else:
            # It's a section
            index = self.tree.indexOfTopLevelItem(item)
            self.tree.takeTopLevelItem(index)
        self.sync_tree_to_model()

    # --- FILE I/O ---
    def save_csv(self):
        self.sync_tree_to_model()
        fn, _ = QFileDialog.getSaveFileName(self, "Save", "", "CSV (*.csv)")
        if fn:
            try:
                with open(fn, 'w', newline='', encoding='utf-8') as f:
                    w = csv.writer(f)
                    md = self.metadata
                    w.writerow(["META", md['school'], md['test'], md['class'], md['subject'], md['time'], md['marks']])
                    for s in self.sections:
                        w.writerow(["SEC", s['name'], s['desc'], s['marks_per_q'], s['attempt_count']])
                        for q in s['questions']:
                            row = ["Q", q['type'], q['text']]
                            if q['type'] == "MCQ": row += q.get('options', [])
                            elif q['type'] == "Match Columns": 
                                row += ["|".join(q.get('col_a', [])), "|".join(q.get('col_b', []))]
                            w.writerow(row)
                QMessageBox.information(self, "Saved", "Progress saved successfully.")
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
                            # Ideally update metadata here, for now keeping simple
                            pass
                        elif row[0] == "SEC":
                            curr_sec = {
                                "name": row[1], "desc": row[2], 
                                "marks_per_q": int(row[3]), "attempt_count": int(row[4]),
                                "total_marks": int(row[3])*int(row[4]), "questions": []
                            }
                            self.sections.append(curr_sec)
                        elif row[0] == "Q" and curr_sec:
                            q = {"type": row[1], "text": row[2]}
                            if q['type'] == "MCQ": q['options'] = row[3:]
                            elif q['type'] == "Match Columns": 
                                q['col_a'] = row[3].split('|')
                                q['col_b'] = row[4].split('|')
                            curr_sec['questions'].append(q)
                self.rebuild_tree()
            except Exception as e:
                print(e)

    # --- PDF EXPORT (UPDATED FOR FONT/SPACING) ---
    def process_text(self, text):
        if not HAS_URDU_LIB or not text: return text, None, False
        is_urdu = bool(re.search('[\u0600-\u06FF]', text))
        if is_urdu:
            reshaped = arabic_reshaper.reshape(text)
            bidi = get_display(reshaped)
            prop = fm.FontProperties(fname=self.urdu_font_path) if self.urdu_font_path else None
            return bidi, prop, True
        return text, None, False

    def export_pdf(self):
        self.sync_tree_to_model()
        fn, _ = QFileDialog.getSaveFileName(self, "Export PDF", f"{self.metadata['subject']}_Exam.pdf", "PDF (*.pdf)")
        if not fn: return

        try:
            # === UPDATED CONSTANTS ===
            PAGE_W, PAGE_H = 8.27, 11.69
            MARGIN_X, MARGIN_TOP, MARGIN_BTM = 0.5, 0.5, 0.5
            CONTENT_W = PAGE_W - (2 * MARGIN_X)
            
            # Increased Font Sizes by 1-2px approx (1 pt ~ 1.3px, keeping logical scale)
            FS_HEADER, FS_SUB, FS_BODY = 18, 14, 12 
            
            # Reduced Line Height (0.25 -> 0.21)
            LH = 0.21 

            with PdfPages(fn) as pdf:
                def new_page():
                    fig = plt.figure(figsize=(PAGE_W, PAGE_H))
                    ax = fig.add_axes([0, 0, 1, 1]) 
                    ax.set_xlim(0, PAGE_W); ax.set_ylim(0, PAGE_H); ax.axis('off')
                    return fig, ax

                def draw_text(ax, x, y, txt, fs=12, weight='normal', align='left', v_align='baseline', force_rtl=True):
                    final_txt, font_prop, is_urdu = self.process_text(txt)
                    eff_x, eff_align = x, align
                    eff_fs = fs

                    if is_urdu:
                        weight = 'normal' # Urdu fonts usually don't support matplotlib bold weights well
                        eff_fs += 1 # Increase Urdu font slightly more for readability
                        if force_rtl:
                            if align == 'left': eff_align, eff_x = 'right', PAGE_W - x
                            elif align == 'right': eff_align, eff_x = 'left', PAGE_W - x

                    kwargs = {'fontsize': eff_fs, 'fontweight': weight, 'ha': eff_align, 'va': v_align}
                    if font_prop: kwargs['fontproperties'] = font_prop
                    ax.text(eff_x, y, final_txt, **kwargs)
                    return is_urdu

                fig, ax = new_page()
                cursor_y = PAGE_H - MARGIN_TOP

                # --- HEADER ---
                H_BOX_H = 2.0
                p_box = FancyBboxPatch((MARGIN_X, cursor_y - H_BOX_H), CONTENT_W, H_BOX_H,
                                       boxstyle="round,pad=0.1", ec="black", fc="white", lw=2)
                ax.add_patch(p_box)
                
                cx = PAGE_W / 2
                draw_text(ax, cx, cursor_y - 0.3, self.metadata['school'], FS_HEADER, 'bold', 'center')
                draw_text(ax, cx, cursor_y - 0.6, self.metadata['test'], FS_SUB, 'bold', 'center')
                
                rule_y = cursor_y - 0.8
                ax.plot([MARGIN_X + 0.2, PAGE_W - MARGIN_X - 0.2], [rule_y, rule_y], color='black', lw=1)
                
                meta_y = rule_y - 0.3
                draw_text(ax, MARGIN_X + 0.2, meta_y, f"Class: {self.metadata['class']}", FS_BODY)
                draw_text(ax, cx + 0.5, meta_y, f"Time: {self.metadata['time']}", FS_BODY)
                
                meta_y -= 0.25
                draw_text(ax, MARGIN_X + 0.2, meta_y, f"Subject: {self.metadata['subject']}", FS_BODY)
                draw_text(ax, cx + 0.5, meta_y, f"Marks: {self.metadata['marks']}", FS_BODY)

                name_y = meta_y - 0.35
                draw_text(ax, MARGIN_X + 0.2, name_y, "Name: __________________________", FS_BODY)
                draw_text(ax, cx + 0.5, name_y, "Roll No: ____________", FS_BODY)
                
                cursor_y -= (H_BOX_H + 0.3)

                # --- SECTIONS ---
                for sec in self.sections:
                    if cursor_y < MARGIN_BTM + 1.0:
                        pdf.savefig(fig); plt.close(); fig, ax = new_page(); cursor_y = PAGE_H - MARGIN_TOP
                    
                    # Section Header
                    SEC_H = 0.4
                    p_sec = FancyBboxPatch((MARGIN_X, cursor_y - SEC_H), CONTENT_W, SEC_H,
                                           boxstyle="round,pad=0.05", ec="black", fc="#ecf0f1", lw=1)
                    ax.add_patch(p_sec)
                    
                    sy = cursor_y - 0.25
                    title = f"{sec['name']}   {sec['desc']}"
                    marks = f"({sec['marks_per_q']} x {sec['attempt_count']} = {sec['total_marks']})"
                    
                    draw_text(ax, MARGIN_X + 0.1, sy, title, FS_BODY, 'bold')
                    draw_text(ax, PAGE_W - MARGIN_X - 0.1, sy, marks, FS_BODY, 'bold', 'right')
                    
                    cursor_y -= (SEC_H + 0.2)

                    # Questions
                    for idx, q in enumerate(sec['questions']):
                        q_num = f"{idx+1}."
                        lines = textwrap.wrap(q['text'], width=80) # Slightly reduced width for larger font
                        
                        # Estimate Height
                        req_h = (len(lines) * LH) + 0.15
                        if q['type'] == "MCQ": req_h += 0.8
                        if q['type'] == "Match Columns": req_h += 1.5
                        
                        if cursor_y - req_h < MARGIN_BTM:
                            pdf.savefig(fig); plt.close(); fig, ax = new_page(); cursor_y = PAGE_H - MARGIN_TOP

                        # Render Text
                        _, _, is_urdu_q = self.process_text(q['text'])
                        
                        if is_urdu_q:
                            draw_text(ax, MARGIN_X, cursor_y, q_num, FS_BODY, 'bold', 'left', 'top', False)
                            anchor = PAGE_W - MARGIN_X - 0.1
                            for ln in lines:
                                draw_text(ax, anchor, cursor_y, ln, FS_BODY, 'normal', 'right', 'top', False)
                                cursor_y -= LH
                        else:
                            draw_text(ax, MARGIN_X, cursor_y, q_num, FS_BODY, 'bold', 'left', 'top', False)
                            anchor = MARGIN_X + 0.5
                            for ln in lines:
                                draw_text(ax, anchor, cursor_y, ln, FS_BODY, 'normal', 'left', 'top', False)
                                cursor_y -= LH
                        
                        # Render Options
                        cursor_y -= 0.1
                        if q['type'] == "MCQ":
                            opts = q.get('options', [])
                            opt_y = cursor_y
                            if is_urdu_q:
                                # Urdu Layout (Right aligned)
                                if len(opts)>0: draw_text(ax, anchor, opt_y, f"{opts[0]} (a)", FS_BODY, 'normal', 'right', 'baseline', False)
                                if len(opts)>2: draw_text(ax, anchor, opt_y-LH, f"{opts[2]} (c)", FS_BODY, 'normal', 'right', 'baseline', False)
                                if len(opts)>1: draw_text(ax, anchor-3.5, opt_y, f"{opts[1]} (b)", FS_BODY, 'normal', 'right', 'baseline', False)
                                if len(opts)>3: draw_text(ax, anchor-3.5, opt_y-LH, f"{opts[3]} (d)", FS_BODY, 'normal', 'right', 'baseline', False)
                                cursor_y = opt_y - (2*LH) - 0.15
                            else:
                                # English Layout
                                if len(opts)>0: draw_text(ax, anchor, opt_y, f"(a) {opts[0]}", FS_BODY)
                                if len(opts)>1: draw_text(ax, anchor+3.5, opt_y, f"(b) {opts[1]}", FS_BODY)
                                opt_y -= LH
                                if len(opts)>2: draw_text(ax, anchor, opt_y, f"(c) {opts[2]}", FS_BODY)
                                if len(opts)>3: draw_text(ax, anchor+3.5, opt_y, f"(d) {opts[3]}", FS_BODY)
                                cursor_y = opt_y - 0.2

                        elif q['type'] == "Match Columns":
                            col_a, col_b = q.get('col_a', []), q.get('col_b', [])
                            draw_text(ax, MARGIN_X+1, cursor_y, "Column A", FS_BODY, 'bold')
                            draw_text(ax, MARGIN_X+4, cursor_y, "Column B", FS_BODY, 'bold')
                            cursor_y -= LH
                            for i in range(max(len(col_a), len(col_b))):
                                if i < len(col_a): draw_text(ax, MARGIN_X+1, cursor_y, col_a[i], FS_BODY)
                                if i < len(col_b): draw_text(ax, MARGIN_X+4, cursor_y, col_b[i], FS_BODY)
                                cursor_y -= LH
                            cursor_y -= 0.2

                pdf.savefig(fig)
                plt.close()
            
            QMessageBox.information(self, "Success", f"PDF Generated:\n{fn}")
            try: os.startfile(fn)
            except: pass

        except Exception as e:
            QMessageBox.critical(self, "PDF Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    setup = SetupDetailsDialog()
    if setup.exec():
        window = ExamGeneratorApp(setup.data)
        window.show()
        sys.exit(app.exec())
