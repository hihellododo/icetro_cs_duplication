import sys
import pandas as pd
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QFileDialog, QLabel, QCalendarWidget, QDateEdit,
                            QComboBox, QSpinBox, QTableView, QMessageBox)
from PyQt6.QtCore import Qt, QDate, QAbstractTableModel, QModelIndex
from PyQt6.QtGui import QColor
from core import Controller

# 데이터프레임을 PyQt의 테이블 뷰(QTableView)에서 표시하기 위한 모델 클래스 정의
class DataFrameModel(QAbstractTableModel):
    def __init__(self, data, parent=None):
        super().__init__(parent)
        # 판다스로 읽어온 데이터 저장
        self._data = data

    def rowCount(self, parent=QModelIndex()):
        '''
        # 데이터프레임의 행 개수를 반환
        '''
        return self._data.shape[0]

    def columnCount(self, parent=QModelIndex()):
        '''
        # 데이터프레임의 열 개수를 반환
        '''
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        # 데이터 표시
        if role == Qt.ItemDataRole.DisplayRole:
            # 셀의 값을 문자열로 반환
            value = self._data.iloc[index.row(), index.column()]
            return str(value)
        
        # 배경색 강조
        if role == Qt.ItemDataRole.BackgroundRole:
            # 3개 이상 중복인 경우 배경색 적용
            if index.column() == 0 and self._data.iloc[index.row(), 0] > 2:
                return QColor('#FFFFCC')
        return None

    def headerData(self, section, orientation, role):
        if role == Qt.ItemDataRole.DisplayRole:
            if orientation == Qt.Orientation.Horizontal:
                # 열 헤더 이름 반환
                return str(self._data.columns[section])
            else:
                # 행 번호 반환
                return str(self._data.index[section])
        return None

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("제품 재불량 분석 시스템")
        self.setGeometry(100, 100, 1200, 800)
        
        # UI 초기화 함수 호출
        self.init_ui()
        # 원본 데이터프레임 저장 변수
        self.df = None
        # 처리된 데이터프레임 저장 변수
        self.processed_df = None

    def init_ui(self):
        """
        UI를 초기화하고 위젯을 배치하는 함수
        """
        main_widget = QWidget()
        main_layout = QVBoxLayout() # 최상위 레이아웃 (세로 배치)

        # 파일 선택 영역
        file_layout = QHBoxLayout() 
        self.file_btn = QPushButton("엑셀 파일 선택") # 파일 선택 버튼 생성
        self.file_btn.clicked.connect(self.load_file) # 버튼 클릭 시 파일 로드 함수 호출
        self.file_label = QLabel("선택된 파일: 없음") # 파일 경로 표시 라벨 생성
        file_layout.addWidget(self.file_btn)
        file_layout.addWidget(self.file_label)

        # 날짜 선택 영역 (달력 위젯)
        date_layout = QHBoxLayout()
        self.start_date = QDateEdit(calendarPopup=True) # 시작 날짜 선택 위젯 (달력 팝업 포함)
        self.start_date.setDate(QDate.currentDate().addMonths(-3)) # 기본값: 현재 날짜 -3개월
        self.end_date = QDateEdit(calendarPopup=True) # 종료 날짜 선택 위젯 (달력 팝업 포함)
        self.end_date.setDate(QDate.currentDate()) # 기본값: 현재 날짜
        date_layout.addWidget(QLabel("시작일:"))
        date_layout.addWidget(self.start_date)
        date_layout.addWidget(QLabel("종료일:"))
        date_layout.addWidget(self.end_date)

        # 컨트롤 영역 (정렬 기준, 중복 개수, 실행 버튼)
        control_layout = QHBoxLayout()
        self.sort_combo = QComboBox() # 정렬 기준 드롭다운 메뉴 생성
        self.sort_combo.addItems(["접수일시", "수리완료일자"]) # 드롭다운 항목 추가
        self.max_duplicate = QSpinBox() # 최대 중복 개수 입력 스핀박스 생성
        self.max_duplicate.setRange(2, 100) # 입력 범위 설정 
        self.max_duplicate.setValue(2) # 기본값 설정
        self.run_btn = QPushButton("실행")  # 실행 버튼 추가
        self.run_btn.clicked.connect(self.process_data)
        control_layout.addWidget(QLabel("정렬 기준:"))
        control_layout.addWidget(self.sort_combo)
        control_layout.addWidget(QLabel("최소 중복 개수:"))
        control_layout.addWidget(self.max_duplicate)
        control_layout.addWidget(self.run_btn)  # 실행 버튼 배치

        # 결과 테이블 뷰 (데이터 표시 영역)
        self.table_view = QTableView()

        # 결과 저장 버튼 생성 및 이벤트 연결
        self.save_btn = QPushButton("결과 저장")
        self.save_btn.clicked.connect(self.save_result)

        # 레이아웃 조립
        main_layout.addLayout(file_layout)       # 파일 선택 레이아웃 추가
        main_layout.addLayout(date_layout)      # 날짜 선택 레이아웃 추가
        main_layout.addLayout(control_layout)   # 정렬/중복 설정 레이아웃 추가
        main_layout.addWidget(self.table_view)  # 결과 테이블 추가
        main_layout.addWidget(self.save_btn)    # 저장 버튼 추가

        main_widget.setLayout(main_layout)      # 최상위 레이아웃 설정
        self.setCentralWidget(main_widget)      # 메인 윈도우에 위젯 설정

    def keyPressEvent(self, event):
        # ESC 키를 누르면 창 닫기
        if event.key() == Qt.Key.Key_Escape:
            self.close()

    def essential_col(self):
        """
        필수 컬럼 항목 확인 
        """
        return ['접수번호', '수리결과', '주소1', '모델코드', 'Serial No.', '접수일시', '수리완료일자', ]

    def load_file(self):
        """
        파일 선택 및 데이터 로드
        """
        fname, _ = QFileDialog.getOpenFileName(self, '파일 선택', '', 'Excel Files (*.xlsx)')
        if fname:
            self.file_label.setText(f"선택된 파일: {fname}")
            try:
                self.df = pd.read_excel(fname)  # 데이터 로드만 수행
                self.table_view.setModel(None)  # 기존 테이블 초기화
            except Exception as e:
                QMessageBox.critical(self, "오류", f"파일 읽기 실패:\n{str(e)}")

    def process_data(self):
        # 만약 파일 로드가 안된 경우 
        if self.df is None:
            QMessageBox.warning(self, "경고", "파일을 먼저 선택해주세요.")
            return
        
        # 필수 컬럼명이 포함된 데이터인지 확인 
        for col in self.essential_col():
            if col not in self.df.columns:
                QMessageBox.critical(self, "오류", f"원본 데이터에 '{col}' 항목이 없습니다.")
                return

        try:
            self.controller = Controller(df=self.df,
                                         start_date=self.start_date,
                                         end_date=self.end_date,
                                         selected_sort_option=self.sort_combo.currentText(),
                                         duplicate_value=self.max_duplicate.value())
            final_df = self.controller.main()

            # 5. 테이블 뷰 업데이트
            model = DataFrameModel(final_df)
            self.table_view.setModel(model)
            self.processed_df = final_df  # 저장용 데이터프레임 업데이트

        except Exception as e:
            QMessageBox.critical(self, "오류", f"데이터 처리 중 오류 발생:\n{str(e)}")

    def save_result(self):
        """
        데이터 저장
        """
        if self.processed_df is not None:
            fname, _ = QFileDialog.getSaveFileName(self, '파일 저장', '', 'Excel Files (*.xlsx)')
            if fname:
                try:
                    with pd.ExcelWriter(fname, engine='openpyxl') as writer:
                        self.processed_df.to_excel(writer, index=False)
                    QMessageBox.information(self, "저장 완료", "파일이 성공적으로 저장되었습니다.")
                except Exception as e:
                    QMessageBox.critical(self, "저장 오류", f"파일 저장 실패:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
