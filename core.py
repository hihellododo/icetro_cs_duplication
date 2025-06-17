import pandas as pd

class Controller():
    def __init__(self, df, start_date, end_date, selected_sort_option, duplicate_value):
        self.df = df
        self.start_date = start_date
        self.end_date = end_date
        self.selected_sort_option = selected_sort_option
        self.duplicate_value = duplicate_value


    def main(self):
        # 1. 데이터 전처리
        df = self.df.copy()
        # 접수 번호 중복 건 제거 
        df.drop_duplicates(subset='접수번호', inplace=True, ignore_index=True)
        # 수리결과에서 접수취소와 자재판매완료 데이터 제거 
        df = df[~df['수리결과'].isin(['접수취소', '자재판매완료'])]
        ## 본사 입고 접수 내역 삭제
        df = df[df['주소1'] != '인천광역시 서구 검단로54번길 7']

        # 2. 날짜 필터링
        start_date = self.start_date.date().toString("yyyy-MM-dd")
        end_date = self.end_date.date().toString("yyyy-MM-dd")
        date_col = '접수일시' if self.selected_sort_option == '접수일시' else '수리완료일자'
        df = df[(df[date_col] >= start_date) & (df[date_col] <= end_date)]

        # 3. 중복 분석
        # 주소1 -> 모델코드 -> Serial No. 순서대로 중복 접수 건 확인 
        duplicate_check_col = ['주소1', '모델코드', 'Serial No.']
        new_df = df[df.duplicated(subset=duplicate_check_col, keep=False)]
        ## 데이터 정렬
        new_df = new_df.sort_values(
            by=['주소1', '모델코드', '접수번호'],
            ignore_index=True
        )

        # 4. 중복 개수 계산
        # 데이터 정렬 기준으로 중복 개수 세기
        duplicate_counts = new_df.groupby(duplicate_check_col, as_index=False).size()
        # 중복건 걸러낸 데이터와 합쳐 출력하기
        final_df = pd.merge(new_df, duplicate_counts, on=duplicate_check_col)  
        # 고유 번호 부여하기 (중복 건수 별로 개수 세기)
        final_df['고유번호']=final_df.groupby(duplicate_check_col, as_index=False).ngroup() + 1
        # 'size' 마지막 열을 맨 앞으로 이동 후 '중복 개수' 이름으로 변경 
        final_df = final_df[list(final_df.columns[-2:]) + list(final_df.columns[:-2])]
        final_df.rename(columns={'size': '중복 개수'}, inplace=True)
        final_df = final_df[final_df['중복 개수'] >= self.duplicate_value]
        return final_df