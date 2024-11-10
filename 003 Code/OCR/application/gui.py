import os
import json
import pandas as pd
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QLabel, 
                             QFileDialog, QProgressBar, QTextEdit, QListWidget, QDialog, QListWidgetItem, QHBoxLayout, QSizePolicy)
from PyQt5.QtGui import QIcon, QColor, QPalette, QDragEnterEvent, QDropEvent
from PyQt5.QtCore import Qt, QTimer, QMimeData
from gpt_service import extract_table_from_text, get_ai_columns
from utils import normalize_column_name
from file_service import calculate_file_hash, manage_cache
from api import ocr_document
import hashlib

class DropArea(QWidget):
    def __init__(self, parent, file_type):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.file_type = file_type
        self.files = []

        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        self.label = QLabel(f"{file_type} 파일 업로드")
        self.label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.label)

        self.list_widget = QListWidget()
        self.list_widget.setStyleSheet("color: white;")
        self.layout.addWidget(self.list_widget)

        self.setMinimumHeight(200)
        self.setMaximumHeight(400) 

        self.update_border_style()

    def update_border_style(self):
        """파일 유무에 따라 스타일 업데이트"""
        if self.files:
            # 파일이 있을 때 (실선)
            self.setStyleSheet(f"""
                DropArea {{
                    background-color: #000000; 
                    border: 2px solid #FFFFFF;  /* 실선으로 변경 */
                    border-radius: 10px;
                    padding: 0px;            
                    margin: 0px;
                    background-position: center;  /* 배경 위치 설정 */
                    background-repeat: no-repeat;
                    background-size: cover;
                }}
            """)
            self.list_widget.setStyleSheet("""
                QListWidget {
                    background-color: rgba(29, 29, 29, 0.4);
                    border: 2px solid #FFFFFF;  /* 실선으로 변경 */
                    border-radius: 10px;
                    padding: 2px;            
                    margin-bottom: 20px;     
                }
                QListWidget::item {
                    background-color: #303030;
                    height: 40px;
                }
                QListWidget::item:selected {
                    background-color: #2F3A5F;  /* 선택된 항목 배경색 */
                }
            """)

        else:
            # 파일이 없을 때 (점선)
            self.setStyleSheet(f"""
                DropArea {{
                    background-color: #000000; 
                    border: 2px dashed #000000; /* 점선으로 변경 */
                    border-radius: 10px;
                    padding: 0px;
                    margin: 0px;
                    background-position: center;  /* 배경 위치 설정 */
                    background-repeat: no-repeat;
                    background-size: cover;
                }}
            """)
            self.list_widget.setStyleSheet("""
                QListWidget {
                    background-color: rgba(29, 29, 29, 0.4);
                    border: 2px dashed #3C3C3C; /* 점선으로 변경 */
                    padding: 20px;            
                    margin-bottom: 20px;     
                }
                QListWidget::item {
                    border: 1px solid white;
                }
                QListWidget::item:selected {
                    background-color: #3C3C3C;
                }
            """)

    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile().endswith('.pdf')]
        self.files.extend(files)
        self.update_list()
        self.update_border_style()

    def update_list(self):
        self.list_widget.clear()
        for file in self.files:
            item = QListWidgetItem(os.path.basename(file))
            item.setForeground(QColor("white"))
            self.list_widget.addItem(item)

class DropAreaImage1(DropArea):
    def __init__(self, parent, file_type):
        super().__init__(parent, file_type)

        # 배경 이미지를 표시할 QLabel 추가
        self.background_label = QLabel(self)  # 먼저 정의
        image_path = os.path.join(os.path.dirname(__file__), "Icon", "DropAreaImage1.png").replace("\\", "/")
        
        self.background_label.setStyleSheet(f"""
            QLabel {{
                background-image: url({image_path});
                background-repeat: no-repeat;
                background-position: center;
                background-size: cover;
            }}
        """)
        self.background_label.setGeometry(self.rect()) 
        self.background_label.lower()

        self.background_label.resize(self.size())
        self.background_label.setScaledContents(True)

        # 업데이트 스타일 호출을 이동
        self.update_border_style()

    def resizeEvent(self, event):
        """DropArea 크기가 변경될 때 배경 이미지를 재조정"""
        super().resizeEvent(event)
        self.background_label.resize(self.size())

    def update_border_style(self):
        """파일 유무에 따라 스타일 업데이트 및 배경 이미지 처리"""
        super().update_border_style()
        # background_label이 존재하는지 확인
        if hasattr(self, "background_label"):
            if self.files:
                self.background_label.hide()  # 파일이 있으면 배경 이미지 숨기기
            else:
                self.background_label.show()  # 파일이 없으면 배경 이미지 표시


class DropAreaImage2(DropArea):
    def __init__(self, parent, file_type):
        super().__init__(parent, file_type)

        # 배경 이미지를 표시할 QLabel 추가
        self.background_label = QLabel(self)  # 먼저 정의
        image_path = os.path.join(os.path.dirname(__file__), "Icon", "DropAreaImage2.png").replace("\\", "/")
        
        self.background_label.setStyleSheet(f"""
            QLabel {{
                background-image: url({image_path});
                background-repeat: no-repeat;
                background-position: center;
                background-size: cover;
            }}
        """)
        self.background_label.setGeometry(self.rect())
        self.background_label.lower()

        self.background_label.resize(self.size())
        self.background_label.setScaledContents(True)

        # 업데이트 스타일 호출을 이동
        self.update_border_style()

    def resizeEvent(self, event):
        """DropArea 크기가 변경될 때 배경 이미지를 재조정"""
        super().resizeEvent(event)
        self.background_label.resize(self.size())

    def update_border_style(self):
        """파일 유무에 따라 스타일 업데이트 및 배경 이미지 처리"""
        super().update_border_style()
        # background_label이 존재하는지 확인
        if hasattr(self, "background_label"):
            if self.files:
                self.background_label.hide()  # 파일이 있으면 배경 이미지 숨기기
            else:
                self.background_label.show()  # 파일이 없으면 배경 이미지 표시




class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.existing_df = None
        
        self.setWindowTitle("스마트 지류to데이터 시스템")
        self.setGeometry(100, 100, 800, 600)
        icon_path = os.path.join(os.path.dirname(__file__), "Icon", "Logo.png").replace("\\", "/")
        
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: #000000;
                background-image: url({icon_path});
                background-repeat: no-repeat;
                background-position: center;
                background-attachment: fixed;
            }}
            QPushButton {{
                background-color: #5D5F5E;
                color: white;
                border: none;
                padding: 10px;
                margin: 5px;
                font-size: 14px;
                border-radius: 5px;
                width: 120px; 
            }}
            QLabel {{
                color: white;
                font-size: 16px;
            }}
        """)


        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        self.setup_ui()

    def setup_ui(self):
        self.top_spacer = QWidget()
        self.top_spacer.setFixedHeight(20)
        self.layout.addWidget(self.top_spacer)

        # 메시지 라벨 (기본 숨김)
        self.message_label = QLabel("")
        self.message_label.setAlignment(Qt.AlignRight)
        # self.message_label.setStyleSheet("""
        #     QLabel {
        #         color: white;  /* 텍스트 색상  + 적용안됨*/
        #         background-color: rgba(0, 0, 0, 0.8);  
        #         border: 2px solid #4687FF;  
        #         border-radius: 10px; 
        #         padding: 10px;  
        #         font-size: 16px;  
        #     }
        # """)
        self.message_label.setVisible(False)
        self.layout.addWidget(self.message_label)

        self.original_drop_area = DropAreaImage1(self, "원본 서류 PDF 파일 업로드")
        self.user_drop_area = DropAreaImage2(self, "설문 결과 PDF 파일 업로드")

        drop_layout = QHBoxLayout()
        drop_layout.addWidget(self.original_drop_area)
        drop_layout.addWidget(self.user_drop_area)
        self.layout.addLayout(drop_layout)

        progress_layout = QVBoxLayout()
        progress_bar_with_percent = QVBoxLayout()

        self.percent_label = QLabel("0%")
        self.percent_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 14px;
                padding: 0px;
                margin: 0px;
            }
        """)
        self.percent_label.setAlignment(Qt.AlignRight) 
        self.percent_label.setFixedHeight(20) 
        self.percent_label.setVisible(False) 
        progress_bar_with_percent.addWidget(self.percent_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #5D5F5E;
                border-radius: 5px;
                background-color: #1C1C1C;
            }
            QProgressBar::chunk {
                background-color: #5D5F5E;
                border-radius: 5px;
            }
        """)
        self.progress_bar.setTextVisible(False) 
        self.progress_bar.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed) 
        self.progress_bar.setFixedHeight(20) 
        progress_bar_with_percent.addWidget(self.progress_bar)

        progress_layout.addLayout(progress_bar_with_percent)
        self.layout.addLayout(progress_layout)

        button_layout = QHBoxLayout()
        button_layout.addStretch(1)
        self.process_button = QPushButton("OCR 진행")
        self.process_button.clicked.connect(self.process_files)
        button_layout.addWidget(self.process_button)

        self.save_button = QPushButton("엑셀 저장")
        self.save_button.clicked.connect(self.save_results_to_excel)
        button_layout.addWidget(self.save_button)

        self.layout.addLayout(button_layout)

        self.check_existing_excel()
        
    def check_existing_excel(self):
        dialog = QDialog(self)
        dialog.setWindowTitle('엑셀 파일 유무')
        dialog.setFixedSize(600, 300)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #2E3B4E;
                color: white;
                font-size: 16px;
            }
            QLabel {
                color: white;
                font-size: 18px;
            }
            QPushButton {
                background-color: #4687FF;
                color: white;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #3A76E3;
            }
        """)

        layout = QVBoxLayout()
        label = QLabel('기존 엑셀 파일이 있으십니까?')
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        button_layout = QHBoxLayout()
        yes_button = QPushButton('Yes')
        no_button = QPushButton('No')

        button_layout.addStretch()
        button_layout.addWidget(yes_button)
        button_layout.addWidget(no_button)
        button_layout.addStretch()

        layout.addLayout(button_layout)
        dialog.setLayout(layout)

        # 버튼 클릭 이벤트 설정
        yes_button.clicked.connect(lambda: (dialog.accept(), self.load_existing_excel()))
        no_button.clicked.connect(lambda: (dialog.reject(), self.show_message("엑셀 파일이 없으시군요, 자동으로 테이블을 생성할게요")))
        
        dialog.exec_()
        
    # def show_message(self, message, duration=3000):
    #     """OCR 진행, 엑셀 저장 또는 오류 시 메시지 표시"""
    #     # 메시지 텍스트 설정
    #     self.message_label.setText(message)
    #     self.message_label.setVisible(True)

    #     # CSS 스타일을 우선순위 높게 설정
    #     self.message_label.setStyleSheet("""
    #         QLabel {
    #             color: white;  /* 텍스트 색상 */
    #             background-color: rgba(0, 0, 0, 0.8);  /* 반투명 배경 */
    #             border: 2px solid #4687FF;  /* 테두리 색상 */
    #             border-radius: 10px;  /* 테두리 둥글기 */
    #             padding: 10px;  /* 안쪽 여백 */
    #             font-size: 16px;  /* 글꼴 크기 */
    #             min-width: 350px;  /* 최소 너비 */
    #             max-width: 500px;  /* 최대 너비 */
    #             text-align: center;  /* 텍스트 가운데 정렬 */
    #         }
    #     """)

    #     # 메시지 라벨 크기를 텍스트에 맞게 조정
    #     self.message_label.adjustSize()

    #     # 메시지 라벨을 창의 우측 상단에 배치
    #     margin = 20  # 우측 및 상단 여백
    #     self.message_label.move(
    #         self.width() - self.message_label.width() - margin,  # 우측 정렬
    #         margin  # 상단 여백
    #     )

    #     # 일정 시간 후 메시지 숨기기
    #     QTimer.singleShot(duration, lambda: self.message_label.setVisible(False))

    def resizeEvent(self, event):
        """창 크기 변경 시 메시지 위치를 재계산"""
        super().resizeEvent(event)
        if self.message_label.isVisible():
            margin = 20
            self.message_label.move(
                self.width() - self.message_label.width() - margin,  # 우측 정렬
                margin  # 상단 여백
            )
            
            
    def load_existing_excel(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            try:
                self.existing_df = pd.read_excel(file_path)
                self.existing_df = self.existing_df.loc[:, ~self.existing_df.columns.str.startswith('Unnamed')]
                self.existing_df = self.existing_df.dropna(axis=1, how='all')        
                self.show_status(f"Loaded existing Excel file: {file_path}")
            except Exception as e:
                self.show_status(f"Error loading Excel file: {e}", error=True)
                
    def process_files(self):
        CACHE_DIR = 'cache'

        # 원본 및 사용자 파일 확인
        if not self.original_drop_area.files:
            self.show_message("Error: Original file not selected", error=True)
            return
        if not self.user_drop_area.files:
            self.show_message("Error: User files not selected", error=True)
            return

        self.percent_label.setVisible(True)

        # 총 파일 개수 계산
        total_files = len(self.user_drop_area.files) + 1

        self.progress_bar.setValue(0)
        self.progress_bar.setMaximum(total_files)
        self.update_progress_label(0, total_files)
        QApplication.processEvents()

        # 원본 파일 처리
        original_file = self.original_drop_area.files[0]
        original_file_hash = calculate_file_hash(original_file)
        original_cache_file = manage_cache(original_file_hash, original_file, CACHE_DIR, ocr_document, 'input')

        try:
            if original_cache_file:
                with open(original_cache_file, "r") as f:
                    original_ocr_data = json.load(f)
            else:
                original_ocr_data = ocr_document(original_file, 'input')
                if original_ocr_data:
                    cache_path = os.path.join(CACHE_DIR, f"{original_file_hash}.json")
                    with open(cache_path, "w") as f:
                        json.dump(original_ocr_data, f)
        except IOError as e:
            self.show_message(f"Error processing original file: {e}", error=True)
            return

        original_text = " ".join([page.get('text', '') for page in original_ocr_data.get('pages', [])])

        # 기존 엑셀 파일 존재 여부에 따른 컬럼 결정
        if self.existing_df is not None:
            self.columns = list(self.existing_df.columns)
        else:
            self.columns = get_ai_columns(original_text)
            if not self.columns:
                self.show_message("Error: Failed to extract columns from AI", error=True)
                return

        self.progress_bar.setValue(1)
        self.update_progress_label(1, total_files)
        QApplication.processEvents()

        # 사용자 파일 처리
        self.results = []
        for i, user_file in enumerate(self.user_drop_area.files):
            user_file_hash = calculate_file_hash(user_file)
            user_cache_file = manage_cache(user_file_hash, user_file, CACHE_DIR, ocr_document, 'user_data')

            try:
                if user_cache_file:
                    with open(user_cache_file, "r") as f:
                        user_ocr_data = json.load(f)
                else:
                    user_ocr_data = ocr_document(user_file, 'user_data')
                    if user_ocr_data:
                        cache_path = os.path.join(CACHE_DIR, f"{user_file_hash}.json")
                        with open(cache_path, "w") as f:
                            json.dump(user_ocr_data, f)
            except IOError as e:
                self.show_message(f"Error processing file {i + 1}: {e}", error=True)
                continue

            if user_ocr_data:
                user_text = " ".join([page.get('text', '') for page in user_ocr_data.get('pages', [])])
                df = extract_table_from_text(user_text, self.columns)
                if not df.empty:
                    self.results.append(df)

            self.progress_bar.setValue(i + 2)
            self.update_progress_label(i + 2, total_files)
            QApplication.processEvents()

        self.show_message("파일 OCR 처리가 완료되었습니다")
        self.update_progress_label(total_files, total_files)
        QApplication.processEvents()

    def save_results_to_excel(self):
        if not self.results:
            self.show_message("Error: No results to save", error=True)
            return

        combined_df = pd.concat(self.results, ignore_index=True)

        if self.existing_df is not None:
            combined_df.columns = [col.strip() for col in combined_df.columns]
            self.existing_df.columns = [col.strip() for col in self.existing_df.columns]

            if list(self.existing_df.columns) == list(combined_df.columns):
                final_df = pd.concat([self.existing_df, combined_df], ignore_index=True)
            else:
                self.show_message("Error: Column headers do not match between existing Excel and new data", error=True)
                return
        else:
            final_df = combined_df

        output_file, _ = QFileDialog.getSaveFileName(self, "Save results as", "", "Excel files (*.xlsx)")
        if output_file:
            try:
                final_df.to_excel(output_file, index=False)
                self.show_message("파일이 엑셀로 변환되어 저장되었습니다")
            except Exception as e:
                self.show_message(f"Failed to save results: {e}", error=True)



    def resizeEvent(self, event):
        """창 크기 변경 시 메시지 위치를 재계산"""
        super().resizeEvent(event)
        if self.message_label.isVisible():
            margin = 20
            self.message_label.move(
                self.width() - self.message_label.width() - margin,  # 우측 정렬
                margin  # 상단 여백
            )


    def update_progress_label(self, current, total=None, error=False):
        """퍼센트 값을 업데이트하고 레이블 표시"""
        if error:
            self.percent_label.setText(current)
            self.percent_label.setStyleSheet("color: red; font-size: 18px;")
        else:
            progress_percent = int((current / total) * 100)
            self.percent_label.setText(f"{progress_percent}%")
            self.percent_label.setStyleSheet("color: white; font-size: 18px;")

    # 메시지 표시 함수 (수정)
    def show_message(self, message, duration=3000):
        """OCR 진행, 엑셀 저장 또는 오류 시 메시지 표시"""
        self.message_label.setText(message)
        self.message_label.setVisible(True)

        # 메시지를 오른쪽 상단으로 고정
        self.message_label.resize(self.message_label.sizeHint())
        margin = 20
        self.message_label.move(self.width() - self.message_label.width() - margin, margin)

        # 일정 시간 후 메시지 숨기기
        QTimer.singleShot(duration, lambda: self.message_label.setVisible(False))


    def update_message_position(self):
        """메시지 라벨의 위치를 오른쪽 상단으로 업데이트"""
        self.message_label.move(self.width() - self.message_label.width() - 20, 20)
        

    def show_status(self, message, error=False):
        color = "red" if error else "white"
        formatted_message = f'<font color="{color}">{message}</font>'
        self.show_message(formatted_message)


if __name__ == "__main__":
    app = QApplication([])
    window = OCRApp()
    window.show()
    app.exec_()