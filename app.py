"""
KITT 엑셀 변환 시스템
Streamlit 기반 웹 애플리케이션
"""
import streamlit as st
import pandas as pd
from excel_processor import ExcelProcessor


# 페이지 설정
st.set_page_config(
    page_title="송장 관리 AI",
    page_icon="📦",
    layout="wide"
)

# 세션 상태 초기화
if 'processor' not in st.session_state:
    st.session_state.processor = ExcelProcessor()
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'output_file' not in st.session_state:
    st.session_state.output_file = None


def main():
    """메인 애플리케이션"""

    # 헤더
    st.title("📦 다들림푸드 송장 관리 AI")
    st.markdown("---")

    # 사이드바 - 사용 안내
    with st.sidebar:
        st.header("📖 사용 방법")
        st.markdown("""
        1. 엑셀 파일(.xlsx, .xls) 업로드
        2. 원본 데이터 확인
        3. 변환 버튼 클릭
        4. 변환된 파일 다운로드
        """)

    # 메인 컨텐츠
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("1️⃣ 파일 업로드")
        uploaded_file = st.file_uploader(
            "엑셀 파일을 선택하세요",
            type=['xlsx', 'xls'],
            help="최대 200MB까지 업로드 가능"
        )

    # 파일이 업로드되면 처리
    if uploaded_file is not None:
        try:
            # 원본 데이터 읽기
            df_original = pd.read_excel(uploaded_file)

            with col1:
                st.success(f"✅ 파일 업로드 완료: **{uploaded_file.name}**")
                st.info(f"📊 데이터 크기: {df_original.shape[0]}행 × {df_original.shape[1]}열")

                # 원본 데이터 미리보기
                with st.expander("🔍 원본 데이터 미리보기", expanded=True):
                    st.dataframe(
                        df_original.head(10),
                        use_container_width=True,
                        height=300
                    )

                # 변환 버튼
                st.markdown("---")
                transform_button = st.button(
                    "🔄 변환 실행",
                    type="primary",
                    use_container_width=True
                )

                if transform_button:
                    with st.spinner("⏳ 변환 중..."):
                        # 파일 포인터 리셋
                        uploaded_file.seek(0)

                        # 변환 실행
                        transformed_df, output_file = st.session_state.processor.process_excel(
                            uploaded_file
                        )

                        # 세션에 저장
                        st.session_state.processed_data = transformed_df
                        st.session_state.output_file = output_file

                        st.success("✅ 변환 완료!")
                        st.balloons()

            # 변환 결과 표시
            with col2:
                if st.session_state.processed_data is not None:
                    st.subheader("2️⃣ 변환 결과")

                    # 변환 요약
                    st.success(f"✅ 변환 완료: {st.session_state.processed_data.shape[0]}행 × {st.session_state.processed_data.shape[1]}열")

                    # 변환된 데이터 미리보기
                    with st.expander("🔍 변환된 데이터 미리보기", expanded=True):
                        st.dataframe(
                            st.session_state.processed_data.head(10),
                            use_container_width=True,
                            height=300
                        )

                    # 다운로드 버튼
                    st.markdown("---")
                    st.subheader("3️⃣ 파일 다운로드")

                    # 다운로드 파일명 생성
                    original_name = uploaded_file.name.rsplit('.', 1)[0]
                    download_filename = f"{original_name}_송장업로드용.xlsx"

                    st.download_button(
                        label="📥 변환된 파일 다운로드",
                        data=st.session_state.output_file,
                        file_name=download_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

                    st.info(f"💾 파일명: `{download_filename}`")

        except Exception as e:
            st.error(f"❌ 오류 발생: {str(e)}")
            st.exception(e)

    else:
        # 파일 업로드 대기 중
        pass

    # 하단 푸터
    st.markdown("---")
    st.caption("🏢 회사 내부용 | 송장 관리 AI v1.0")


if __name__ == "__main__":
    main()
