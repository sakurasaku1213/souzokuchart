import sys
import uuid
import csv
import os
import tempfile

try:
    import openpyxl
    from openpyxl.drawing.image import Image as OpenpyxlImage
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("警告: openpyxlライブラリが見つかりません。Excelエクスポート機能は無効になります。`pip install openpyxl`でインストールしてください。")


from PyQt5.QtWidgets import (
    QMainWindow, QVBoxLayout, QHBoxLayout, QWidget, QLabel, QLineEdit,
    QComboBox, QPushButton, QDateEdit, QFormLayout, QApplication,
    QListWidget, QListWidgetItem, QMessageBox,
    QGraphicsView, QGraphicsScene, QGraphicsRectItem, QGraphicsTextItem, QGraphicsLineItem,
    QGraphicsItem,
    QFileDialog,
    QAction, QMenuBar, QDockWidget,
    QTextEdit, QTabWidget, QStyle
)
from PyQt5.QtGui import QPainter, QPen, QBrush, QColor, QImage, QKeySequence, QIcon
from PyQt5.QtCore import QDate, Qt, QPointF, QRectF, QSize

from srz_kun.core.data_model import Person
from srz_kun.core.project_file import save_project, load_project

# --- Constants for Diagram Items ---
NODE_WIDTH = 160
NODE_HEIGHT = 70
V_SPACING = 60
H_SPACING = 60

# --- PersonNodeItem Class Definition ---
class PersonNodeItem(QGraphicsRectItem):
    def __init__(self, person_id: str, display_text: str, x: float, y: float, is_deceased=False, parent_item=None):
        super().__init__(x, y, NODE_WIDTH, NODE_HEIGHT, parent_item)
        self.person_id = person_id
        self.connected_lines = []

        pen_color = Qt.black
        pen_width = 2 if is_deceased else 1.5
        pen = QPen(pen_color, pen_width)
        self.setPen(pen)

        brush_color = QColor("#e0e0e0") if is_deceased else QColor("#f8f8f8")
        self.setBrush(QBrush(brush_color))

        self.setFlag(QGraphicsItem.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.ItemSendsGeometryChanges, True)

        self.text_item = QGraphicsTextItem(display_text, self)
        text_rect = self.text_item.boundingRect()
        text_x = (NODE_WIDTH - text_rect.width()) / 2
        text_y = (NODE_HEIGHT - text_rect.height()) / 2
        self.text_item.setPos(text_x, text_y)

    def add_line(self, line_item: QGraphicsLineItem, other_node: 'PersonNodeItem', is_this_node_start: bool):
        self.connected_lines.append({
            'line': line_item,
            'other_node': other_node,
            'is_this_node_start': is_this_node_start
        })

    def get_center_point(self) -> QPointF:
        return QPointF(self.pos().x() + NODE_WIDTH / 2, self.pos().y() + NODE_HEIGHT / 2)

    def itemChange(self, change, value):
        if change == QGraphicsItem.ItemPositionHasChanged and self.scene():
            new_this_node_center = QPointF(value.x() + NODE_WIDTH / 2, value.y() + NODE_HEIGHT / 2)
            for line_info in self.connected_lines:
                line_item = line_info['line']
                other_node_item = line_info['other_node']
                is_this_node_start = line_info['is_this_node_start']
                other_node_center = other_node_item.get_center_point()
                if is_this_node_start:
                    line_item.setLine(new_this_node_center.x(), new_this_node_center.y(),
                                      other_node_center.x(), other_node_center.y())
                else:
                    line_item.setLine(other_node_center.x(), other_node_center.y(),
                                      new_this_node_center.x(), new_this_node_center.y())
        return super().itemChange(change, value)

    def __del__(self):
        pass


class MainWindowUI(QMainWindow):
    APP_VERSION = "1.0.0"
    APP_NAME = "相続関係図くん"
    DEFAULT_CHECKLIST_ITEMS = [
        "戸籍謄本（被相続人：出生から死亡まで全て）",
        "除籍謄本・改製原戸籍（被相続人）",
        "住民票の除票（被相続人：本籍地記載あり）",
        "戸籍謄本（相続人全員：現在戸籍）",
        "住民票（相続人全員：本籍地記載あり）",
        "印鑑証明書（相続人全員）",
        "遺産分割協議書（作成した場合）",
        "固定資産評価証明書（不動産がある場合）",
        "預貯金通帳・残高証明書（該当金融機関分）",
        "有価証券・会員権等の証書",
        "生命保険証書（受取人が被相続人または相続財産となっている場合）",
        "借入金残高証明書・契約書（負債がある場合）"
    ]
    INHERITANCE_PATTERNS_GUIDE = {
        "-- パターンを選択 --": "<em>ここに選択した相続パターンの基本的な説明、法定相続分、注意点、作成すべき図の範囲の目安などが表示されます。</em>",
        "ケース1: 配偶者と子": "<h3>ケース1: 配偶者と子</h3>" \
                           "<b>説明:</b> 被相続人の配偶者とすべての子（実子・養子）が相続人となります。<br>" \
                           "<b>法定相続分:</b> 配偶者 1/2, 子 全体で 1/2（子の間で均等割）。<br>" \
                           "<b>注意点:</b><ul><li>子が先に死亡している場合、その子の子（被相続人の孫）が代襲相続することがあります。</li><li>非嫡出子の相続分は嫡出子と同じです。</li></ul>" \
                           "<b>作成範囲の目安:</b> 被相続人、配偶者、すべての子（および代襲相続する孫など）を記載します。",
        "ケース2: 子のみ": "<h3>ケース2: 子のみ</h3>" \
                         "<b>説明:</b> 被相続人に配偶者がいない（または既に死亡）場合で、子がいる場合、すべての子が相続人となります。<br>" \
                         "<b>法定相続分:</b> 子がすべてを相続（子の間で均等割）。<br>" \
                         "<b>注意点:</b><ul><li>子が先に死亡している場合、その子の子（被相続人の孫）が代襲相続することがあります。</li></ul>" \
                         "<b>作成範囲の目安:</b> 被相続人、すべての子（および代襲相続する孫など）を記載します。",
        "ケース3: 配偶者と直系尊属": "<h3>ケース3: 配偶者と直系尊属（父母など）</h3>" \
                               "<b>説明:</b> 被相続人に子がいない（または代襲相続する孫などもいない）場合で、配偶者と直系尊属（父母、祖父母など）がいる場合、これらが相続人となります。<br>" \
                               "<b>法定相続分:</b> 配偶者 2/3, 直系尊属 全体で 1/3（父母ともに健在なら各1/6）。<br>" \
                               "<b>注意点:</b><ul><li>より親等の近い直系尊属が優先されます（例：父母健在なら祖父母は相続しない）。</li></ul>" \
                               "<b>作成範囲の目安:</b> 被相続人、配偶者、該当する直系尊属を記載します。",
        "ケース4: 直系尊属のみ": "<h3>ケース4: 直系尊属（父母など）のみ</h3>" \
                              "<b>説明:</b> 被相続人に配偶者も子もいない場合、直系尊属が相続人となります。<br>" \
                              "<b>法定相続分:</b> 直系尊属がすべてを相続。<br>" \
                              "<b>作成範囲の目安:</b> 被相続人、該当する直系尊属を記載します。",
        "ケース5: 配偶者と兄弟姉妹": "<h3>ケース5: 配偶者と兄弟姉妹</h3>" \
                                 "<b>説明:</b> 被相続人に子も直系尊属もいない場合で、配偶者と兄弟姉妹がいる場合、これらが相続人となります。<br>" \
                                 "<b>法定相続分:</b> 配偶者 3/4, 兄弟姉妹 全体で 1/4。<br>" \
                                 "<b>注意点:</b><ul><li>兄弟姉妹が先に死亡している場合、その子（被相続人の甥姪）が代襲相続することがあります。</li><li>半血の兄弟姉妹の相続分は全血の兄弟姉妹の1/2です。</li></ul>" \
                                 "<b>作成範囲の目安:</b> 被相続人、配偶者、兄弟姉妹（および代襲相続する甥姪など）を記載します。",
        "ケース6: 兄弟姉妹のみ": "<h3>ケース6: 兄弟姉妹のみ</h3>" \
                              "<b>説明:</b> 被相続人に配偶者も子も直系尊属もいない場合、兄弟姉妹が相続人となります。<br>" \
                              "<b>法定相続分:</b> 兄弟姉妹がすべてを相続。<br>" \
                              "<b>作成範囲の目安:</b> 被相続人、兄弟姉妹（および代襲相続する甥姪など）を記載します。"
    }

    def __init__(self):
        super().__init__()
        self.current_project_filepath = None
        self.loaded_layout_positions = {}
        self.setWindowTitle(f"{self.APP_NAME} Ver. {self.APP_VERSION} - Untitled")
        self.setGeometry(100, 100, 1200, 750)

        self.people_list = []
        self.diagram_items = {}
        self.diagram_lines = []

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_h_layout = QHBoxLayout(self.central_widget)

        self.left_panel_widget = QWidget()
        self.left_panel_v_layout = QVBoxLayout(self.left_panel_widget)
        self.template_label = QLabel("家族構成テンプレート:")
        self.left_panel_v_layout.addWidget(self.template_label)
        self.family_template_combo = QComboBox()
        self.family_template_combo.addItems([
            "-- テンプレートを選択 --", "被相続人・配偶者・子供1人", "被相続人・子供2人"
        ])
        self.family_template_combo.currentIndexChanged.connect(self._on_family_template_selected)
        self.left_panel_v_layout.addWidget(self.family_template_combo)
        self.left_panel_v_layout.addSpacing(20)

        self.form_layout = QFormLayout()
        self.name_input = QLineEdit()
        self.form_layout.addRow(QLabel("氏名:"), self.name_input)
        self.perm_domicile_input = QLineEdit()
        self.form_layout.addRow(QLabel("本籍地:"), self.perm_domicile_input)
        self.address_input = QLineEdit()
        self.form_layout.addRow(QLabel("住所:"), self.address_input)
        self.relation_input = QComboBox()
        self.relation_input.addItems(["被相続人", "妻", "夫", "長男", "長女", "養子", "父", "母", "子", "その他"])
        self.relation_input.currentTextChanged.connect(self._on_relationship_changed) # Connect signal
        self.form_layout.addRow(QLabel("続柄:"), self.relation_input)
        self.status_input = QComboBox()
        self.status_input.addItems(["生存", "死亡"])
        self.status_input.currentIndexChanged.connect(self._on_is_alive_status_changed)
        self.form_layout.addRow(QLabel("生死の別:"), self.status_input)
        self.waiver_input = QComboBox()
        self.waiver_input.addItems(["なし", "あり"])
        self.form_layout.addRow(QLabel("相続放棄の有無:"), self.waiver_input)
        self.dob_input = QDateEdit()
        self.dob_input.setDisplayFormat("yyyy-MM-dd")
        self.dob_input.setDate(QDate.currentDate().addYears(-30))
        self.form_layout.addRow(QLabel("生年月日:"), self.dob_input)
        self.dod_input = QDateEdit()
        self.dod_input.setDisplayFormat("yyyy-MM-dd")
        self.dod_input.setNullable(True)
        self.dod_input.setSpecialValueText(" ")
        self.dod_input.setDate(self.dod_input.minimumDate()) # Ensure it starts null
        self.form_layout.addRow(QLabel("死亡年月日:"), self.dod_input)
        self.left_panel_v_layout.addLayout(self.form_layout)

        self.add_person_button = QPushButton("人物を追加")
        self.add_person_button.clicked.connect(self._on_add_person_clicked)
        self.left_panel_v_layout.addWidget(self.add_person_button, alignment=Qt.AlignCenter)

        self.export_csv_button = QPushButton("CSVエクスポート")
        self.export_csv_button.clicked.connect(self._on_export_csv_clicked)
        self.left_panel_v_layout.addWidget(self.export_csv_button, alignment=Qt.AlignCenter)

        self.export_excel_button = QPushButton("Excelエクスポート (図)")
        self.export_excel_button.clicked.connect(self._on_export_excel_clicked)
        if not OPENPYXL_AVAILABLE:
            self.export_excel_button.setDisabled(True)
            self.export_excel_button.setToolTip("openpyxlライブラリが必要です。pip install openpyxl")
        self.left_panel_v_layout.addWidget(self.export_excel_button, alignment=Qt.AlignCenter)

        self.left_panel_v_layout.addSpacing(10)
        self.left_panel_v_layout.addStretch()
        self.people_list_widget = QListWidget()
        self.left_panel_v_layout.addWidget(self.people_list_widget)
        self.main_h_layout.addWidget(self.left_panel_widget, 1)

        self.scene = QGraphicsScene()
        self.diagram_view = QGraphicsView(self.scene)
        self.diagram_view.setRenderHint(QPainter.Antialiasing)
        self.diagram_view.setDragMode(QGraphicsView.ScrollHandDrag)
        self.main_h_layout.addWidget(self.diagram_view, 3)

        self.support_dock_widget = QDockWidget("サポート", self)
        self.support_tab_widget = QTabWidget()
        self.checklist_widget = QListWidget()
        for text in self.DEFAULT_CHECKLIST_ITEMS:
            item = QListWidgetItem(text, self.checklist_widget)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
        self.support_tab_widget.addTab(self.checklist_widget, "チェックリスト")
        guide_page_widget = QWidget()
        guide_layout = QVBoxLayout(guide_page_widget)
        guide_label = QLabel("相続パターンを選択してください:")
        self.pattern_guide_combo = QComboBox()
        self.pattern_guide_combo.addItems(self.INHERITANCE_PATTERNS_GUIDE.keys())
        self.pattern_guide_display = QTextEdit()
        self.pattern_guide_display.setReadOnly(True)
        guide_layout.addWidget(guide_label)
        guide_layout.addWidget(self.pattern_guide_combo)
        guide_layout.addWidget(self.pattern_guide_display)
        guide_page_widget.setLayout(guide_layout)
        self.support_tab_widget.addTab(guide_page_widget, "相続パターンガイド")
        self.support_dock_widget.setWidget(self.support_tab_widget)
        self.addDockWidget(Qt.RightDockWidgetArea, self.support_dock_widget)
        self.support_dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.pattern_guide_combo.currentTextChanged.connect(self._on_pattern_guide_selected)
        self._on_pattern_guide_selected(self.pattern_guide_combo.currentText())

        self._create_menus()
        self._on_is_alive_status_changed()
        self._on_relationship_changed(self.relation_input.currentText()) # Call for initial state
        self._redraw_diagram()

        self.family_template_combo.setToolTip("一般的な家族構成をリストに自動入力します。現在のリスト内容はクリアされます。")
        self.add_person_button.setToolTip("入力された情報をリストに追加し、図に反映します。")
        self.people_list_widget.setToolTip("追加された人物の一覧です。")
        self.diagram_view.setToolTip("相続関係図の表示エリアです。ノードをドラッグして位置を調整できます。")
        self.export_csv_button.setToolTip("現在リストにある人物情報をCSVファイルとしてエクスポートします。")
        self.export_excel_button.setToolTip("現在の相続関係図を画像としてExcelファイルにエクスポートします。")
        self.checklist_widget.setToolTip("必要な書類や手続きの確認用チェックリストです。状態はプロジェクトと共に保存されます。")
        self.pattern_guide_combo.setToolTip("代表的な相続パターンを選択して説明を読みます。")
        self.pattern_guide_display.setToolTip("選択された相続パターンの説明、注意点、作成範囲の目安などが表示されます。")
        self.support_dock_widget.setToolTip("関連作業のチェックリストと相続パターンのガイドです。")

        try:
            self.add_person_button.setIcon(self.style().standardIcon(QStyle.SP_ListPlus))
        except:
            print("Note: QStyle.SP_ListPlus icon not available or failed to set.")
        self.export_csv_button.setIcon(self.style().standardIcon(QStyle.SP_ArrowUp))
        self.export_excel_button.setIcon(self.style().standardIcon(QStyle.SP_ArrowUp))

    def _on_pattern_guide_selected(self, pattern_name):
        guide_html = self.INHERITANCE_PATTERNS_GUIDE.get(pattern_name, "<p>情報が見つかりません。</p>")
        self.pattern_guide_display.setHtml(guide_html)

    def _on_relationship_changed(self, relationship_text: str):
        if relationship_text == "被相続人":
            self.status_input.setCurrentText("死亡")
            self.status_input.setEnabled(False)
            self.waiver_input.setCurrentText("なし")
            self.waiver_input.setEnabled(False)
        else:
            self.status_input.setEnabled(True)
            self.waiver_input.setEnabled(True)
        self._on_is_alive_status_changed() # Ensure DoD input state is updated

    def _gather_project_data(self) -> dict:
        people_data = []
        for p in self.people_list:
            person_dict = {
                "id": p.id, "name": p.name,
                "relationship_to_deceased": p.relationship_to_deceased,
                "date_of_birth": p.date_of_birth, "date_of_death": p.date_of_death,
                "permanent_domicile": p.permanent_domicile, "address": p.address,
                "is_alive": p.is_alive, "waived_inheritance": p.waived_inheritance
            }
            people_data.append(person_dict)

        diagram_layout_data = {
            pid: {'x': item.pos().x(), 'y': item.pos().y()}
            for pid, item in self.diagram_items.items() if item
        }

        checklist_data = []
        for i in range(self.checklist_widget.count()):
            item = self.checklist_widget.item(i)
            checklist_data.append({'text': item.text(), 'checked': item.checkState() == Qt.Checked})

        return {
            "application_version": self.APP_VERSION,
            "people": people_data,
            "diagram_layout": diagram_layout_data,
            "checklist": checklist_data
        }

    def _on_save_project(self):
        if not self.current_project_filepath:
            self._on_save_as_project()
            return

        if not self.people_list and not any(self.diagram_items.values()) and self.current_project_filepath is None :
             reply = QMessageBox.question(self, "空のプロジェクトの保存",
                                         "人物リストは空です。空のプロジェクトとして保存しますか？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
             if reply == QMessageBox.No:
                return

        project_data = self._gather_project_data()
        try:
            if save_project(project_data, self.current_project_filepath):
                QMessageBox.information(self, "保存完了", f"プロジェクトが保存されました:\n{self.current_project_filepath}")
            else:
                QMessageBox.critical(self, "保存失敗", "プロジェクトの保存中にエラーが発生しました。")
        except Exception as e:
            QMessageBox.critical(self, "保存失敗", f"予期せぬエラーが発生しました: {e}")
            print(f"Error during save_project call: {e}")

    def _on_save_as_project(self):
        if not self.people_list and not any(self.diagram_items.values()):
             reply = QMessageBox.question(self, "空のプロジェクトの保存",
                                         "人物リストは空です。空のプロジェクトとして保存しますか？",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
             if reply == QMessageBox.No:
                return

        current_name = os.path.basename(self.current_project_filepath) if self.current_project_filepath else "untitled.srk"

        options = QFileDialog.Options()
        filepath, _ = QFileDialog.getSaveFileName(self, "名前を付けてプロジェクトを保存", current_name,
                                                  f"{self.APP_NAME}プロジェクト (*.srk);;All Files (*)", options=options)
        if filepath:
            self.current_project_filepath = filepath
            project_data = self._gather_project_data()
            try:
                if save_project(project_data, self.current_project_filepath):
                    QMessageBox.information(self, "保存完了", f"プロジェクトが保存されました:\n{self.current_project_filepath}")
                    self.setWindowTitle(f"{self.APP_NAME} Ver. {self.APP_VERSION} - {os.path.basename(self.current_project_filepath)}")
                else:
                    QMessageBox.critical(self, "保存失敗", "プロジェクトの保存中にエラーが発生しました。")
            except Exception as e:
                QMessageBox.critical(self, "保存失敗", f"予期せぬエラーが発生しました: {e}")
                print(f"Error during save_project call (Save As): {e}")

    def _on_load_project(self):
        if self.people_list or self.current_project_filepath:
            reply = QMessageBox.question(self, "確認",
                                         "現在のプロジェクトの変更は破棄されます。新しいプロジェクトを読み込みますか？\n（変更がある場合は先に保存してください）",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                return

        options = QFileDialog.Options()
        filepath, _ = QFileDialog.getOpenFileName(self, "プロジェクトファイルを開く", "",
                                                  f"{self.APP_NAME}プロジェクト (*.srk);;All Files (*)", options=options)
        if filepath:
            loaded_data = load_project(filepath)
            if loaded_data is None or not isinstance(loaded_data, dict):
                QMessageBox.critical(self, "エラー", f"プロジェクトファイルの読み込みに失敗しました。\nファイル形式が不正か、破損している可能性があります。\n\n詳細: {filepath}")
                return

            self._clear_all_people_data()

            diagram_layout_from_file = loaded_data.get("diagram_layout", {})
            if not isinstance(diagram_layout_from_file, dict):
                print(f"警告: 'diagram_layout' データが不正な形式です。無視されます。")
                self.loaded_layout_positions = {}
            else:
                self.loaded_layout_positions = diagram_layout_from_file

            loaded_checklist_items = loaded_data.get("checklist", [])
            if not isinstance(loaded_checklist_items, list):
                print(f"警告: 'checklist' データが不正な形式です。無視されます。")
                loaded_checklist_items = []

            if self.checklist_widget.count() != len(self.DEFAULT_CHECKLIST_ITEMS):
                self.checklist_widget.clear()
                for text in self.DEFAULT_CHECKLIST_ITEMS:
                    item = QListWidgetItem(text, self.checklist_widget)
                    item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                    item.setCheckState(Qt.Unchecked)

            for i in range(self.checklist_widget.count()):
                ui_item = self.checklist_widget.item(i)
                ui_item_text = ui_item.text()
                found_in_loaded = next((loaded_item for loaded_item in loaded_checklist_items if loaded_item.get('text') == ui_item_text), None)
                if found_in_loaded:
                    ui_item.setCheckState(Qt.Checked if found_in_loaded.get('checked', False) else Qt.Unchecked)
                else:
                    ui_item.setCheckState(Qt.Unchecked)

            people_raw_data = loaded_data.get("people", [])
            if not isinstance(people_raw_data, list):
                QMessageBox.critical(self, "エラー", "人物データが不正な形式です。")
                self.loaded_layout_positions = {}
                for i in range(self.checklist_widget.count()):
                    self.checklist_widget.item(i).setCheckState(Qt.Unchecked)
                return

            temp_people_list = []
            for p_data in people_raw_data:
                if not isinstance(p_data, dict):
                    print(f"警告: person data item is not a dictionary: {p_data}. Skipping.")
                    continue
                person = Person(
                    id=p_data.get('id'),
                    name=p_data.get('name', "不明な氏名"),
                    relationship_to_deceased=p_data.get('relationship_to_deceased', "不明"),
                    date_of_birth=p_data.get('date_of_birth', ""),
                    date_of_death=p_data.get('date_of_death'),
                    permanent_domicile=p_data.get('permanent_domicile', ""),
                    address=p_data.get('address', ""),
                    is_alive=p_data.get('is_alive', True),
                    waived_inheritance=p_data.get('waived_inheritance', False)
                )
                temp_people_list.append(person)
                item = QListWidgetItem(str(person))
                item.setData(Qt.UserRole, person.id)
                self.people_list_widget.addItem(item)

            self.people_list = temp_people_list
            self._redraw_diagram()

            self.current_project_filepath = filepath
            self.setWindowTitle(f"{self.APP_NAME} Ver. {self.APP_VERSION} - {os.path.basename(self.current_project_filepath)}")
            QMessageBox.information(self, "成功", f"プロジェクト「{os.path.basename(filepath)}」を読み込みました。")

            self.loaded_layout_positions = {}

    def _redraw_diagram(self):
        self.scene.clear()
        self.diagram_items.clear()
        self.diagram_lines.clear()

        use_loaded_layout = bool(self.loaded_layout_positions)

        if not self.people_list:
            message_item = QGraphicsTextItem("人物リストにデータを追加してください。")
            self.scene.addItem(message_item)
            self.scene.setSceneRect(-100, -50, 200, 100)
            return

        deceased_persons = [p for p in self.people_list if p.relationship_to_deceased == "被相続人"]
        if not deceased_persons:
            self.scene.addItem(QGraphicsTextItem("被相続人が指定されていません。"))
            self.scene.setSceneRect(-100, -50, 200, 100)
            return

        deceased_person = deceased_persons[0]

        spouses = [p for p in self.people_list if p.relationship_to_deceased in ["妻", "夫"] and p.id != deceased_person.id]
        spouse_person = spouses[0] if spouses else None
        children = [p for p in self.people_list if p.relationship_to_deceased in ["長男", "長女", "養子", "子"] and p.id != deceased_person.id and (not spouse_person or p.id != spouse_person.id)]

        auto_current_x = 50
        auto_current_y = 50

        deceased_pos_data = self.loaded_layout_positions.get(deceased_person.id) if use_loaded_layout else None
        deceased_x = deceased_pos_data['x'] if deceased_pos_data else auto_current_x
        deceased_y = deceased_pos_data['y'] if deceased_pos_data else auto_current_y

        deceased_display_text = f"{deceased_person.name}\n(被相続人)"
        if not deceased_person.is_alive and deceased_person.date_of_death: deceased_display_text += f"\n死亡: {deceased_person.date_of_death}"
        elif not deceased_person.is_alive: deceased_display_text += f"\n(死亡)"
        deceased_node = PersonNodeItem(deceased_person.id, deceased_display_text, deceased_x, deceased_y, is_deceased=True)
        self.scene.addItem(deceased_node)
        self.diagram_items[deceased_person.id] = deceased_node

        spouse_node = None
        if spouse_person:
            spouse_pos_data = self.loaded_layout_positions.get(spouse_person.id) if use_loaded_layout else None
            auto_spouse_x = deceased_node.pos().x() + NODE_WIDTH + H_SPACING
            auto_spouse_y = deceased_node.pos().y()
            spouse_x = spouse_pos_data['x'] if spouse_pos_data else auto_spouse_x
            spouse_y = spouse_pos_data['y'] if spouse_pos_data else auto_spouse_y

            spouse_display_text = f"{spouse_person.name}\n({spouse_person.relationship_to_deceased})"
            if not spouse_person.is_alive and spouse_person.date_of_death: spouse_display_text += f"\n死亡: {spouse_person.date_of_death}"
            elif not spouse_person.is_alive: spouse_display_text += f"\n(死亡)"
            spouse_node = PersonNodeItem(spouse_person.id, spouse_display_text, spouse_x, spouse_y)
            self.scene.addItem(spouse_node)
            self.diagram_items[spouse_person.id] = spouse_node

            p1_center = deceased_node.get_center_point()
            p2_center = spouse_node.get_center_point()
            line_between_spouses = QGraphicsLineItem(p1_center.x() + NODE_WIDTH / 2, p1_center.y(),
                                                     p2_center.x() - NODE_WIDTH / 2, p2_center.y())
            self.scene.addItem(line_between_spouses)
            self.diagram_lines.append(line_between_spouses)
            deceased_node.add_line(line_between_spouses, spouse_node, True)
            spouse_node.add_line(line_between_spouses, deceased_node, False)

        if children:
            children_y_level_nodes = deceased_node.pos().y() + NODE_HEIGHT + V_SPACING
            children_bus_line_y = children_y_level_nodes - V_SPACING / 2
            parent_connector_x = deceased_node.get_center_point().x()
            parent_connector_y_start = deceased_node.get_center_point().y() + NODE_HEIGHT / 2
            if spouse_node:
                parent_connector_x = (deceased_node.get_center_point().x() + spouse_node.get_center_point().x()) / 2
                parent_connector_y_start = deceased_node.get_center_point().y()
            main_vertical_line = QGraphicsLineItem(parent_connector_x, parent_connector_y_start,
                                                   parent_connector_x, children_bus_line_y)
            self.scene.addItem(main_vertical_line)
            self.diagram_lines.append(main_vertical_line)
            num_children = len(children)
            total_children_width = num_children * NODE_WIDTH + (num_children - 1) * H_SPACING
            auto_children_start_x = parent_connector_x - total_children_width / 2
            actual_child_centers_x = []
            for i, child_person in enumerate(children):
                child_pos_data = self.loaded_layout_positions.get(child_person.id) if use_loaded_layout else None
                child_x = child_pos_data['x'] if child_pos_data else (auto_children_start_x + i * (NODE_WIDTH + H_SPACING))
                child_y = child_pos_data['y'] if child_pos_data else children_y_level_nodes
                child_display_text = f"{child_person.name}\n({child_person.relationship_to_deceased})"
                if not child_person.is_alive and child_person.date_of_death: child_display_text += f"\n死亡: {child_person.date_of_death}"
                elif not child_person.is_alive: child_display_text += f"\n(死亡)"
                child_node = PersonNodeItem(child_person.id, child_display_text, child_x, child_y)
                self.scene.addItem(child_node)
                self.diagram_items[child_person.id] = child_node
                actual_child_centers_x.append(child_node.get_center_point().x())

            if num_children > 1 and actual_child_centers_x:
                children_horizontal_line = QGraphicsLineItem(min(actual_child_centers_x), children_bus_line_y,
                                                             max(actual_child_centers_x), children_bus_line_y)
                self.scene.addItem(children_horizontal_line)
                self.diagram_lines.append(children_horizontal_line)
                for child_id_idx, child_center_x_val in enumerate(actual_child_centers_x):
                    child_person_obj = children[child_id_idx]
                    child_node_item = self.diagram_items.get(child_person_obj.id)
                    if child_node_item:
                        line_to_child = QGraphicsLineItem(child_center_x_val, children_bus_line_y,
                                                          child_center_x_val, child_node_item.pos().y() )
                        self.scene.addItem(line_to_child)
                        self.diagram_lines.append(line_to_child)
            elif num_children == 1 and actual_child_centers_x:
                child_node_item = self.diagram_items.get(children[0].id)
                if child_node_item:
                    main_vertical_line.setLine(parent_connector_x, parent_connector_y_start,
                                               actual_child_centers_x[0], child_node_item.pos().y())

        if self.scene.items():
            try:
                rect = self.scene.itemsBoundingRect()
                padding = 20
                self.scene.setSceneRect(rect.adjusted(-padding, -padding, padding, padding))
            except Exception as e:
                print(f"Error setting scene rect: {e}")
                min_x, max_x, min_y, max_y = float('inf'), float('-inf'), float('inf'), float('-inf')
                if not self.diagram_items:
                     self.scene.setSceneRect(-200, -100, 400, 200)
                else:
                    for item_id, item_obj in self.diagram_items.items():
                        item_rect = item_obj.sceneBoundingRect()
                        min_x = min(min_x, item_rect.left())
                        max_x = max(max_x, item_rect.right())
                        min_y = min(min_y, item_rect.top())
                        max_y = max(max_y, item_rect.bottom())
                    if min_x != float('inf'):
                        self.scene.setSceneRect(min_x - padding, min_y - padding, (max_x - min_x) + 2*padding, (max_y - min_y) + 2*padding)
                    else:
                        self.scene.setSceneRect(-200, -200, 400, 400)
        else:
            self.scene.setSceneRect(-200, -100, 400, 200)

    def _add_person_to_gui(self, person: Person):
        self.people_list.append(person)
        list_item = QListWidgetItem(str(person))
        list_item.setData(Qt.UserRole, person.id)
        self.people_list_widget.addItem(list_item)
        self._redraw_diagram()

    def _clear_all_people_data(self):
        self.people_list.clear()
        self.people_list_widget.clear()
        self.loaded_layout_positions = {}
        self.current_project_filepath = None
        self.setWindowTitle(f"{self.APP_NAME} Ver. {self.APP_VERSION} - Untitled")
        if hasattr(self, 'checklist_widget'):
            for i in range(self.checklist_widget.count()):
                self.checklist_widget.item(i).setCheckState(Qt.Unchecked)
        self._redraw_diagram()

    def _on_family_template_selected(self, index):
        template_name = self.family_template_combo.itemText(index)
        if index == 0:
            return
        if self.people_list or self.current_project_filepath:
            reply = QMessageBox.question(self, "確認",
                                         "現在のプロジェクトの変更は破棄されます。テンプレートを適用しますか？\n（変更がある場合は先に保存してください）",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.No:
                self.family_template_combo.setCurrentIndex(0)
                return
        self._clear_all_people_data()
        default_dob_deceased = QDate.currentDate().addYears(-70).toString("yyyy-MM-dd")
        default_dod_deceased = QDate.currentDate().addYears(-1).toString("yyyy-MM-dd")
        default_dob_spouse = QDate.currentDate().addYears(-65).toString("yyyy-MM-dd")
        default_dob_child1 = QDate.currentDate().addYears(-40).toString("yyyy-MM-dd")
        default_dob_child2 = QDate.currentDate().addYears(-35).toString("yyyy-MM-dd")
        default_perm_domicile = "東京都千代田区"
        default_address = "同上"
        if template_name == "被相続人・配偶者・子供1人":
            deceased = Person(name="被相続人 例太郎", relationship_to_deceased="被相続人", date_of_birth=default_dob_deceased, permanent_domicile=default_perm_domicile, address=default_address, is_alive=False, waived_inheritance=False, date_of_death=default_dod_deceased)
            self._add_person_to_gui(deceased)
            spouse = Person(name="配偶者 例花子", relationship_to_deceased="妻", date_of_birth=default_dob_spouse, permanent_domicile=default_perm_domicile, address=default_address, is_alive=True, waived_inheritance=False)
            self._add_person_to_gui(spouse)
            child1 = Person(name="子供 例一郎", relationship_to_deceased="長男", date_of_birth=default_dob_child1, permanent_domicile=default_perm_domicile, address=default_address, is_alive=True, waived_inheritance=False)
            self._add_person_to_gui(child1)
        elif template_name == "被相続人・子供2人":
            deceased = Person(name="被相続人 例二郎", relationship_to_deceased="被相続人", date_of_birth=default_dob_deceased, permanent_domicile=default_perm_domicile, address=default_address, is_alive=False, waived_inheritance=False, date_of_death=default_dod_deceased)
            self._add_person_to_gui(deceased)
            child1 = Person(name="子供 例良子", relationship_to_deceased="長女", date_of_birth=default_dob_child1, permanent_domicile=default_perm_domicile, address=default_address, is_alive=True, waived_inheritance=False)
            self._add_person_to_gui(child1)
            child2 = Person(name="子供 例三郎", relationship_to_deceased="長男", date_of_birth=default_dob_child2, permanent_domicile=default_perm_domicile, address=default_address, is_alive=True, waived_inheritance=False)
            self._add_person_to_gui(child2)
        self.family_template_combo.setCurrentIndex(0)

    def _on_is_alive_status_changed(self):
        status_text = self.status_input.currentText()
        if status_text == "生存":
            self.dod_input.setEnabled(False)
            self.dod_input.setDate(self.dod_input.minimumDate()) # Reset to nullable state
            self.dod_input.setSpecialValueText(" ")
        else:
            self.dod_input.setEnabled(True)

    def _on_add_person_clicked(self):
        name = self.name_input.text()
        if not name.strip():
            QMessageBox.warning(self, "入力エラー", "氏名を入力してください。")
            return
        perm_domicile = self.perm_domicile_input.text()
        address = self.address_input.text()
        relation = self.relation_input.currentText()
        status_text = self.status_input.currentText()
        is_alive = True if status_text == "生存" else False

        waiver_text = self.waiver_input.currentText()
        if relation == "被相続人":
            waived_inheritance = False
        else:
            waived_inheritance = True if waiver_text == "あり" else False

        dob_qdate = self.dob_input.date()
        dob_str = dob_qdate.toString("yyyy-MM-dd")

        dod_qdate = self.dod_input.date()
        dod_str = None

        if not is_alive:
            if self.dod_input.isEnabled() and dod_qdate.isValid() and \
               not (self.dod_input.specialValueText() and dod_qdate == self.dod_input.minimumDate()):

                dod_str = dod_qdate.toString("yyyy-MM-dd")

                if dob_qdate.isValid() and dod_qdate < dob_qdate:
                    QMessageBox.warning(self, "入力エラー", "死亡年月日は生年月日より後の日付である必要があります。")
                    return
            elif relation == "被相続人":
                if not (dod_qdate.isValid() and not (self.dod_input.specialValueText() and dod_qdate == self.dod_input.minimumDate())):
                    pass

        new_person = Person(name=name, relationship_to_deceased=relation, date_of_birth=dob_str, permanent_domicile=perm_domicile, address=address, is_alive=is_alive, waived_inheritance=waived_inheritance, date_of_death=dod_str)
        if self.loaded_layout_positions:
            self.loaded_layout_positions = {}
        self._add_person_to_gui(new_person)
        self._clear_input_fields()

    def _clear_input_fields(self):
        self.name_input.clear()
        self.perm_domicile_input.clear()
        self.address_input.clear()
        self.relation_input.setCurrentIndex(0)
        self.status_input.setCurrentIndex(0)
        # self.status_input.setEnabled(True) # This will be handled by _on_relationship_changed
        self.waiver_input.setCurrentIndex(0)
        # self.waiver_input.setEnabled(True) # This will be handled by _on_relationship_changed
        self.dob_input.setDate(QDate.currentDate().addYears(-30))
        self.dod_input.setDate(self.dod_input.minimumDate())
        self.dod_input.setSpecialValueText(" ")
        self.name_input.setFocus()
        self._on_relationship_changed(self.relation_input.currentText())


    def _create_menus(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("ファイル")

        save_action = QAction("保存", self)
        save_action.setShortcut(QKeySequence.Save)
        save_action.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        save_action.triggered.connect(self._on_save_project)
        file_menu.addAction(save_action)

        save_as_action = QAction("名前を付けて保存...", self)
        save_as_action.setShortcut(QKeySequence.SaveAs)
        save_as_action.triggered.connect(self._on_save_as_project)
        file_menu.addAction(save_as_action)

        file_menu.addSeparator()

        load_action = QAction("読み込み...", self)
        load_action.setShortcut(QKeySequence.Open)
        load_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_action.triggered.connect(self._on_load_project)
        file_menu.addAction(load_action)


    def _on_export_csv_clicked(self):
        if not self.people_list:
            QMessageBox.information(self, "エクスポート不可", "エクスポートする人物データがありません。")
            return
        default_filename = "相続関係情報.csv"
        deceased_persons = [p for p in self.people_list if p.relationship_to_deceased == "被相続人"]
        if deceased_persons:
            deceased_name_sanitized = "".join(c if c.isalnum() or c in (' ', '_') else '_' for c in deceased_persons[0].name)
            default_filename = f"{deceased_name_sanitized}_相続関係情報.csv"
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getSaveFileName(self, "CSVエクスポート", default_filename, "CSV Files (*.csv);;All Files (*)", options=options)
        if filePath:
            headers = ["ID", "氏名", "続柄", "生年月日", "死亡年月日", "本籍地", "住所", "生死", "相続放棄"]
            try:
                with open(filePath, 'w', newline='', encoding='utf-8-sig') as csvfile:
                    writer = csv.writer(csvfile)
                    writer.writerow(headers)
                    for person in self.people_list:
                        is_alive_str = "生存" if person.is_alive else "死亡"
                        waived_inheritance_str = "あり" if person.waived_inheritance else "なし"
                        row = [person.id, person.name, person.relationship_to_deceased, person.date_of_birth, person.date_of_death if person.date_of_death else "", person.permanent_domicile, person.address, is_alive_str, waived_inheritance_str]
                        writer.writerow(row)
                QMessageBox.information(self, "エクスポート完了", f"データが正常にエクスポートされました:\n{filePath}")
            except Exception as e:
                QMessageBox.critical(self, "エクスポート失敗", f"CSVファイルのエクスポート中にエラーが発生しました:\n{e}")
                print(f"CSV Export Error: {e}")

    def _on_export_excel_clicked(self):
        if not OPENPYXL_AVAILABLE:
            QMessageBox.warning(self, "機能制限", "Excelエクスポートに必要なライブラリ (openpyxl) が見つかりません。")
            return
        if not self.people_list or not self.scene.items():
            QMessageBox.information(self, "エクスポート不可", "エクスポートする図データがありません。")
            return
        scene_rect = self.scene.itemsBoundingRect()
        if scene_rect.isEmpty():
            QMessageBox.warning(self, "エクスポートエラー", "図の範囲を取得できませんでした。")
            return
        padding = 20
        padded_rect = scene_rect.adjusted(-padding, -padding, padding, padding)
        if padded_rect.width() <= 0 or padded_rect.height() <= 0:
             QMessageBox.warning(self, "エクスポートエラー", f"図のサイズが無効です: 幅={padded_rect.width()}, 高さ={padded_rect.height()}")
             return
        image = QImage(padded_rect.size().toSize(), QImage.Format_ARGB32_Premultiplied)
        image.fill(Qt.white)
        painter = QPainter(image)
        try:
            self.scene.render(painter, QRectF(image.rect()), padded_rect)
        finally:
            painter.end()
        temp_image_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_image_file:
                temp_image_path = temp_image_file.name
            if not image.save(temp_image_path):
                QMessageBox.critical(self, "エクスポート失敗", "一時画像ファイルの保存に失敗しました。")
                if temp_image_path and os.path.exists(temp_image_path): os.remove(temp_image_path)
                return
            default_filename = "相続関係図.xlsx"
            deceased_persons = [p for p in self.people_list if p.relationship_to_deceased == "被相続人"]
            if deceased_persons:
                deceased_name_sanitized = "".join(c if c.isalnum() or c in (' ', '_') else '_' for c in deceased_persons[0].name)
                default_filename = f"{deceased_name_sanitized}_相続関係図.xlsx"
            options = QFileDialog.Options()
            excel_filepath, _ = QFileDialog.getSaveFileName(self, "Excelファイルとして保存", default_filename, "Excel Files (*.xlsx);;All Files (*)", options=options)
            if not excel_filepath:
                if temp_image_path and os.path.exists(temp_image_path): os.remove(temp_image_path)
                return
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "相続関係図"
            img = OpenpyxlImage(temp_image_path)
            sheet.add_image(img, 'A1')
            workbook.save(excel_filepath)
            QMessageBox.information(self, "エクスポート完了", f"図が正常にExcelファイルとしてエクスポートされました:\n{excel_filepath}")
        except Exception as e:
            QMessageBox.critical(self, "エクスポート失敗", f"Excelファイルのエクスポート中にエラーが発生しました:\n{e}")
            print(f"Excel Export Error: {e}")
        finally:
            if temp_image_path and os.path.exists(temp_image_path):
                try:
                    os.remove(temp_image_path)
                except Exception as e_remove:
                    print(f"Error deleting temporary file {temp_image_path}: {e_remove}")

    def load_data(self, project_data):
        print("Direct load_data called (placeholder)...")
        self._clear_all_people_data()
        diagram_layout_from_file = project_data.get("diagram_layout", {})
        if not isinstance(diagram_layout_from_file, dict):
            self.loaded_layout_positions = {}
        else:
            self.loaded_layout_positions = diagram_layout_from_file
        loaded_checklist_items = project_data.get("checklist", [])
        if not isinstance(loaded_checklist_items, list):
            loaded_checklist_items = []
        if self.checklist_widget.count() != len(self.DEFAULT_CHECKLIST_ITEMS):
            self.checklist_widget.clear()
            for text in self.DEFAULT_CHECKLIST_ITEMS:
                item = QListWidgetItem(text, self.checklist_widget)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
        for i in range(self.checklist_widget.count()):
            ui_item = self.checklist_widget.item(i)
            ui_item_text = ui_item.text()
            found_in_loaded = next((loaded_item for loaded_item in loaded_checklist_items if loaded_item.get('text') == ui_item_text), None)
            if found_in_loaded:
                ui_item.setCheckState(Qt.Checked if found_in_loaded.get('checked', False) else Qt.Unchecked)
            else:
                ui_item.setCheckState(Qt.Unchecked)
        people_raw_data = project_data.get("people", [])
        if not isinstance(people_raw_data, list):
            self.loaded_layout_positions = {}
            for i in range(self.checklist_widget.count()):
                self.checklist_widget.item(i).setCheckState(Qt.Unchecked)
            return
        temp_people_list = []
        for p_data in people_raw_data:
            if not isinstance(p_data, dict):
                continue
            person = Person(
                id=p_data.get('id'), name=p_data.get('name'),
                relationship_to_deceased=p_data.get('relationship_to_deceased'),
                date_of_birth=p_data.get('date_of_birth'), date_of_death=p_data.get('date_of_death'),
                permanent_domicile=p_data.get('permanent_domicile'), address=p_data.get('address'),
                is_alive=p_data.get('is_alive'), waived_inheritance=p_data.get('waived_inheritance')
            )
            temp_people_list.append(person)
            item = QListWidgetItem(str(person))
            item.setData(Qt.UserRole, person.id)
            self.people_list_widget.addItem(item)
        self.people_list = temp_people_list
        self._redraw_diagram()
        self.loaded_layout_positions = {}

    def get_data_for_saving(self):
        return self._gather_project_data()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_win = MainWindowUI()
    main_win.show()
    sys.exit(app.exec_())
