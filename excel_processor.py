"""
엑셀 파일 변환 로직
"""
import pandas as pd
from io import BytesIO
from typing import Dict, Tuple


class ExcelProcessor:
    """엑셀 변환 처리 클래스"""

    # 품목명 매핑
    PRODUCT_MAPPING = {
        '참기름350ml': '몽.참',
        '들기름350ml': '몽.들',
        '[사은품]볶음참깨 80g': '깨',
        '[사은품]볶음참깨80g': '깨',  # 공백 없는 버전 대비
        '참기름 350ml': '몽.참',  # 공백 있는 버전
        '들기름 350ml': '몽.들',
        '소문난 참기름350ml': '소.참',
        '소문난참기름350ml': '소.참',  # 공백 없는 버전
        '소문난 들기름350ml': '소.들',
        '소문난들기름350ml': '소.들',  # 공백 없는 버전
    }

    # 제품군별 출력 순서
    PRODUCT_ORDER_MONG = ['몽.참', '몽.들', '깨']  # 몽고메 제품군
    PRODUCT_ORDER_SO = ['소.참', '소.들']  # 소문난 제품군

    def __init__(self):
        """초기화"""
        pass

    def convert_product_name(self, items_dict: Dict[str, int]) -> str:
        """
        품목명과 수량을 변환된 포맷으로 변경

        Args:
            items_dict: {품목명: 수량} 딕셔너리

        Returns:
            변환된 품목명 문자열 (예: "몽.참2,몽.들1,깨" 또는 "소.참1,소.들2")
        """
        # 제품군 판별: 소문난 제품이 있는지 확인
        has_so = False
        for product in items_dict.keys():
            matched = self.PRODUCT_MAPPING.get(product.strip())
            if matched and matched.startswith('소.'):
                has_so = True
                break

        # 제품군에 따라 순서 선택
        if has_so:
            product_order = self.PRODUCT_ORDER_SO
        else:
            product_order = self.PRODUCT_ORDER_MONG

        result = []

        # 고정 순서대로 처리
        for short_name in product_order:
            # items_dict에서 해당 약어에 매칭되는 수량 찾기
            quantity = 0
            for product, qty in items_dict.items():
                matched_short = self.PRODUCT_MAPPING.get(product.strip())
                if matched_short == short_name:
                    quantity = qty
                    break

            # 수량이 0이면 숫자 없이 품목명만, 1 이상이면 숫자 포함
            if quantity == 0:
                result.append(short_name)
            else:
                result.append(f"{short_name}{quantity}")

        return ','.join(result) if result else ''

    def determine_type(self, total_quantity: int) -> str:
        """
        총 수량에 따라 타입 결정

        Args:
            total_quantity: 참기름 + 들기름 + 참깨 총 개수

        Returns:
            "A" 또는 "C"
        """
        return "C" if total_quantity >= 15 else "A"

    def process_excel(self, uploaded_file) -> Tuple[pd.DataFrame, BytesIO]:
        """
        엑셀 파일을 읽어서 변환 규칙을 적용

        Args:
            uploaded_file: Streamlit UploadedFile 객체

        Returns:
            tuple: (변환된 DataFrame, 다운로드용 BytesIO 객체)
        """
        # 1. 엑셀 파일 읽기
        df = pd.read_excel(uploaded_file)

        # 2. 컬럼명 정규화 (인덱스 기반으로 안전하게 처리)
        # 원본 컬럼: 주문, 받는사람, 품목, 물량요청, 물량요청.1, 운임규분, 타입, 수량, 주소.1, 전화번호1, 배송메시지, 송장번호, ...
        # 인덱스로 매핑
        if len(df.columns) >= 11:
            col_mapping = {}
            col_mapping[df.columns[0]] = '주문'  # A열
            col_mapping[df.columns[1]] = '받는사람'  # B열
            col_mapping[df.columns[2]] = '품목'  # C열
            col_mapping[df.columns[3]] = '물량요청'  # D열
            col_mapping[df.columns[4]] = '물량요청_운임포함'  # E열
            # F열(운임규분), G열(타입)은 건너뜀 (원본에 없거나 비어있음)
            if len(df.columns) > 7:
                col_mapping[df.columns[7]] = '수량'  # H열
            if len(df.columns) > 8:
                col_mapping[df.columns[8]] = '주소'  # I열
            if len(df.columns) > 9:
                col_mapping[df.columns[9]] = '전화번호1'  # J열
            if len(df.columns) > 10:
                col_mapping[df.columns[10]] = '배송메시지'  # K열

            df.rename(columns=col_mapping, inplace=True)

        # 3. 받는사람 + 주소 기준으로 그룹화 (순서 보장)
        if '받는사람' not in df.columns or '주소' not in df.columns:
            raise ValueError("필수 컬럼(받는사람, 주소)이 없습니다.")

        # 그룹화를 직접 구현하여 원본 순서 100% 보장
        seen_groups = {}  # {group_key: [행 인덱스들]}
        group_order = []   # 그룹이 처음 등장한 순서

        for idx, row in df.iterrows():
            group_key = str(row['받는사람']) + '|||' + str(row['주소'])

            if group_key not in seen_groups:
                seen_groups[group_key] = []
                group_order.append(group_key)

            seen_groups[group_key].append(idx)

        # 그룹 순서대로 처리
        result_rows = []

        for group_key in group_order:
            row_indices = seen_groups[group_key]
            group_df = df.loc[row_indices]

            # 첫 번째 행의 기본 정보 유지
            first_row = group_df.iloc[0]

            # 품목별 수량 집계
            items_dict = {}
            for _, row in group_df.iterrows():
                product = str(row['품목']).strip() if pd.notna(row['품목']) else ''
                quantity = int(row['수량']) if pd.notna(row['수량']) else 0

                if product:
                    items_dict[product] = items_dict.get(product, 0) + quantity

            # 품목명 변환
            converted_product = self.convert_product_name(items_dict)

            # 총 수량 계산
            total_quantity = sum(items_dict.values())

            # 타입 결정
            item_type = self.determine_type(total_quantity)

            # 새로운 행 생성
            new_row = {
                '받는사람': first_row['받는사람'],
                '품목': converted_product,  # B열: 변환된 품목명
                '불필요항목': '식품',  # C열: 고정값 "식품"
                '불필요항목_2': '박스',  # D열: 고정값 "박스"
                '운임구분': '신용',  # 고정값
                '타입': item_type,
                '수량': 1,  # 통합 후 1로 설정
                '주소': first_row['주소'],
                '전화번호1': first_row['전화번호1'] if '전화번호1' in first_row else '',
                '배송메시지': first_row['배송메시지'] if '배송메시지' in first_row and pd.notna(first_row['배송메시지']) else '',
                '송장번호': '',  # K열: 빈 값
            }

            result_rows.append(new_row)

        # 4. 결과 DataFrame 생성 (이미 순서가 보장됨)
        transformed_df = pd.DataFrame(result_rows)

        # NaN을 빈 문자열로 변경
        transformed_df = transformed_df.fillna('')

        # 5. 컬럼 순서 재배치
        final_columns = [
            '받는사람',
            '품목',  # B열: 변환된 품목명
            '불필요항목',  # C열: 식품
            '불필요항목_2',  # D열: 박스
            '운임구분',
            '타입',
            '수량',
            '주소',
            '전화번호1',
            '배송메시지',
            '송장번호'  # K열
        ]

        # 존재하는 컬럼만 선택 (_order 제외)
        existing_columns = [col for col in final_columns if col in transformed_df.columns]
        transformed_df = transformed_df[existing_columns]

        # 6. Excel 파일로 변환 (메모리에 저장)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            transformed_df.to_excel(writer, index=False, sheet_name='송장업로드용')

        output.seek(0)

        return transformed_df, output

    def get_column_info(self, df: pd.DataFrame) -> dict:
        """
        데이터프레임의 컬럼 정보 반환

        Args:
            df: pandas DataFrame

        Returns:
            dict: 컬럼 이름, 타입, 샘플 값 등
        """
        info = {
            'columns': df.columns.tolist(),
            'dtypes': df.dtypes.to_dict(),
            'shape': df.shape,
            'sample': df.head(3).to_dict('records') if len(df) > 0 else []
        }
        return info
