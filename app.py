import streamlit as st
import winrm
import requests
import pyodbc
import pandas as pd
import math
import socket
import time
from datetime import datetime
from pathlib import Path
import base64

st.title("OMD Remote Dashboard")

# Streamlit의 상태 초기화
if "ip_address" not in st.session_state:
    st.session_state.ip_address = ""
if "is_connected" not in st.session_state:
    st.session_state.is_connected = None  # 초기 연결 상태 없음
if "swv" not in st.session_state:
    st.session_state.swv = "Not Available"  # 초기 SWV 값

# MSTSC 연결 상태를 확인하는 함수
def check_mstsc_connection(ip_address):
    if ip_address:
        try:
            session = winrm.Session(
                f'http://{ip_address}:5985/wsman',
                auth=('a2927\\autologonuser', 'Columbia03'),
                transport='ntlm',
                server_cert_validation='ignore',
                read_timeout_sec=30,
            )
            cmd_query = 'query user'
            result_query = session.run_cmd(cmd_query)
            output_query = result_query.std_out.decode('cp1252')

            if output_query and ("console" in output_query.lower() or "rdp" in output_query.lower()):
                return True
            return False
        except Exception as e:
            st.error(f"Error: {e}")
            return False
    return False

# SWV 값을 가져오는 함수
def fetch_swv(ip_address):
    try:
        mcdb_path = '5300%5CMachine%5CSoftware%20Version'  # URL 인코딩된 경로
        url = f'http://{ip_address}:54123/api/MCDB/Read?mcdbPath={mcdb_path}'
        response = requests.get(url, headers={'Accept': 'application/json'})
        if response.status_code == 200:
            response_data = response.json()
            return response_data.get('Value', 'Value not found')
        else:
            return "Failed to fetch data"
    except requests.exceptions.RequestException as e:
        return f"Error fetching data: {e}"

# 사이드바에서 IP 주소 입력 필드
ip_address = st.sidebar.text_input(
    "Enter Tool's ID", 
    st.session_state.ip_address,
    on_change=lambda: st.session_state.update({"ip_address": ip_address})
)

# 연결 상태 확인 버튼
if st.sidebar.button("Check Connection"):
    if ip_address:
        st.session_state.is_connected = check_mstsc_connection(ip_address)
        st.session_state.swv = fetch_swv(ip_address)  # SWV 값 갱신

# 사이드바에 연결 상태 표시
if st.session_state.is_connected is not None:
    if st.session_state.is_connected:
        st.sidebar.markdown('<p style="color:red;">Someone is connected the tool.</p>', unsafe_allow_html=True)
    else:
        st.sidebar.markdown('<p style="color:green;">No one is connected the tool.</p>', unsafe_allow_html=True)

# 사이드바에 SWV 출력
st.sidebar.text(f"SWV: {st.session_state.swv}")

# 탭 UI 구성
tab1, tab2, tab3 , tab4, tab5= st.tabs(["How to Use", "MCDB", "Run History", "DB Result", "Diagnostiscs"])



# Tab1: 간단한 내용 출력
with tab1:

    st.header("How to use tool connection status")
    
    error1="img/connection_error2.png"
    connect="img/connect_status.png"
    connect1="img/connect_status1.png"
    connect2="img/connect_status2.png"
    how_to_connect="img/how_to_connect.png"
    

    st.markdown(
        """

        1. Insert tool ID in input box in left sidebar and press enter
        2. Click 'Check Connection' button
        3. If tool is free, you can see below message with 'Green'
        4. If tool is occupid from someone, you can see below message with 'Red'
        """)
    st.image(connect)
    
    
    st.markdown(
        """
        
        ### How to solve error ###
        If you see below error message after clicking 'Check Connection' button,
        please follow below guidance. it means, the precondition is not correctly set.
        
        """)
    
    st.image(error1)

    st.markdown(
        """
        
        To check the connection status, you have two preconditions.
        1. You have to connect VPN.
        2. Target tool should be set below setting correctly. 
        
        How to setting
        1. Go to the target tool
        2. Run 'Windows PowerShell'
        3. Type 'winrm quickconfig' and press enter
        4. Type 'y' and press enter
        5. You can see the tool connection status once you complete above steps correctly
        """)
    st.image(how_to_connect)


# Tab2: 레지스트리 값 확인
with tab2:
    st.header("MCDB Checker")
    def fetch_registry_value(api_url, mcdb_path):
        try:
            url = f"{api_url}?mcdbPath={mcdb_path}"
            response = requests.get(url, headers={"Accept": "application/json"})
            if response.status_code == 200:
                data = response.json()
                return data.get("Value", "No Value Found")
            else:
                return f"Error: {response.status_code} - {response.text}"
        except Exception as e:
            return f"Error: {e}"

    
    API_URL = f"http://{ip_address}:54123/api/MCDB/Read"
    mcdb_path = st.text_input("Enter Registry Path:", placeholder="e.g., 5300\\Machine\\Software Version")
    if st.button("See Value"):
        if mcdb_path:
            value = fetch_registry_value(API_URL, mcdb_path)
            st.info(f"Result: {value}")
        else:
            st.warning("Please enter a valid mcdbPath.")

# 포트를 찾는 함수
def find_sql_server_port(ip_address):
    try:
        server_port = 1434
        timeout = 2
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(timeout)
        message = b'\x02'
        sock.sendto(message, (ip_address, server_port))
        response, _ = sock.recvfrom(4096)
        sock.close()
        response_str = response.decode('utf-8', errors='ignore')
        entries = response_str.split(';')
        for i in range(len(entries)):
            if entries[i] == "tcp":
                return entries[i + 1]
    except Exception as e:
        st.error(f"Failed to find port for {ip_address}: {e}")
        return None

# Tab3: NGResults Viewer
with tab3:
    st.title("Run History Dashboard")

    # Tab3 기본 설정값
    TAB3_IP_ADDRESS = ip_address
    TAB3_DATABASE = "NGResults"
    TAB3_USERNAME = "sa"
    TAB3_PASSWORD = "atlas"
    TAB3_ROWS_PER_PAGE = 500

    # 기본 쿼리 설정
    base_query_tab3 = """
    SELECT
        R.RecipeName,
        J.JobName,
        J.RecipePath,
        J.Status,
        J.StartTime,
        J.EndTime,
        J.LoadPort,
        J.OperatorName,
        J.ToolId
    FROM [NGResults].[dbo].[RECIPES] R
    INNER JOIN [NGResults].[dbo].[JOBS] J ON R.KeyId = J.Recipe_KeyId
    ORDER BY J.StartTime DESC
    """

    # 상태 초기화
    if "tab3_total_rows" not in st.session_state:
        st.session_state.tab3_total_rows = 0
    if "tab3_page" not in st.session_state:
        st.session_state.tab3_page = 1
    if "tab3_port" not in st.session_state:
        st.session_state.tab3_port = None

    # SQL Server 연결 및 데이터 조회 함수
    def tab3_get_data(offset, fetch, ip_address, port):
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={ip_address},{port};"
                f"DATABASE={TAB3_DATABASE};"
                f"UID={TAB3_USERNAME};"
                f"PWD={TAB3_PASSWORD};"
            )
            paginated_query = f"""
            {base_query_tab3}
            OFFSET {offset} ROWS FETCH NEXT {fetch} ROWS ONLY
            """
            data = pd.read_sql_query(paginated_query, conn)
            conn.close()
            return data
        except Exception as e:
            st.error(f"Connection failed or query error: {e}")
            return None

    # 조회 버튼 클릭 시 데이터 로드
    if st.button("Load Data", key="tab3_search"):
        progress_bar = st.progress(0)
        st.session_state.tab3_port = find_sql_server_port(TAB3_IP_ADDRESS)

        if st.session_state.tab3_port:
            try:
                for i in range(100):  # 프로그래스 바 애니메이션
                    time.sleep(0.02)
                    progress_bar.progress(i + 1)

                conn = pyodbc.connect(
                    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                    f"SERVER={TAB3_IP_ADDRESS},{st.session_state.tab3_port};"
                    f"DATABASE={TAB3_DATABASE};"
                    f"UID={TAB3_USERNAME};"
                    f"PWD={TAB3_PASSWORD};"
                )
                count_query = """
                SELECT COUNT(*) AS TotalRows
                FROM [NGResults].[dbo].[RECIPES] R
                INNER JOIN [NGResults].[dbo].[JOBS] J ON R.KeyId = J.Recipe_KeyId
                """
                st.session_state.tab3_total_rows = pd.read_sql_query(count_query, conn).iloc[0]["TotalRows"]
                conn.close()
                st.session_state.tab3_page = 1
            except Exception as e:
                st.error(f"Connection failed or query error: {e}")
            finally:
                progress_bar.empty()

    # 데이터 표시
    if st.session_state.tab3_total_rows > 0:
        rows_per_page = TAB3_ROWS_PER_PAGE
        offset = (st.session_state.tab3_page - 1) * rows_per_page
        data = tab3_get_data(offset, rows_per_page, TAB3_IP_ADDRESS, st.session_state.tab3_port)
        st.dataframe(data)

# Tab4: omRESULTS Viewer
with tab4:
    st.header("DB Results Viewer")

    # Tab4 기본 설정값
    TAB4_IP_ADDRESS = ip_address
    TAB4_DATABASE = "omRESULTS"
    TAB4_USERNAME = "sa"
    TAB4_PASSWORD = "atlas"
    TAB4_ROWS_PER_PAGE = 1000

    # 기본 쿼리 설정
    base_query_tab4 = """
    SELECT
        LR.LOT_NAME,
        LR.RECIPE_NAME,
        LR.OPERATOR_NAME,
        LR.TOOL_ID,
        LR.RUN_START_TIME,
        LR.RUN_END_TIME,
        LR.ITERATION,
        M.ORIENTATION,
        M.TEST_NUM,
        M.X_DIE,
        M.Y_DIE,
        AM.X_MISREG,
        AM.Y_MISREG,
        AM.X_TIS,
        AM.Y_TIS
    FROM [omRESULTS].[dbo].[LOT_RUN] LR
    INNER JOIN [omRESULTS].[dbo].[MEASUREMENT] M ON LR.[RUN_ID] = M.[RUN_ID]
    INNER JOIN [omRESULTS].[dbo].[AROL_MEASUREMENT] AM ON M.[MEASUREMENT_ID] = AM.[MEASUREMENT_ID]
    ORDER BY LR.RUN_START_TIME DESC
    """

    # 상태 초기화
    if "tab4_total_rows" not in st.session_state:
        st.session_state.tab4_total_rows = 0
    if "tab4_page" not in st.session_state:
        st.session_state.tab4_page = 1
    if "tab4_port" not in st.session_state:
        st.session_state.tab4_port = None

    # SQL Server 연결 및 데이터 조회 함수
    def tab4_get_data(offset, fetch, ip_address, port):
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={ip_address},{port};"
                f"DATABASE={TAB4_DATABASE};"
                f"UID={TAB4_USERNAME};"
                f"PWD={TAB4_PASSWORD};"
            )
            paginated_query = f"""
            {base_query_tab4}
            OFFSET {offset} ROWS FETCH NEXT {fetch} ROWS ONLY
            """
            data = pd.read_sql_query(paginated_query, conn)
            conn.close()
            return data
        except Exception as e:
            st.error(f"Connection failed or query error: {e}")
            return None

    # 조회 버튼 클릭 시
    if st.button("Load Data", key="tab4_search"):
        progress_bar = st.progress(0)
        st.session_state.tab4_port = find_sql_server_port(TAB4_IP_ADDRESS)

        if st.session_state.tab4_port:
            try:
                for i in range(100):  # 프로그래스 바 애니메이션
                    time.sleep(0.02)
                    progress_bar.progress(i + 1)

                conn = pyodbc.connect(
                    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                    f"SERVER={TAB4_IP_ADDRESS},{st.session_state.tab4_port};"
                    f"DATABASE={TAB4_DATABASE};"
                    f"UID={TAB4_USERNAME};"
                    f"PWD={TAB4_PASSWORD};"
                )
                count_query = """
                SELECT COUNT(*) AS TotalRows
                FROM [omRESULTS].[dbo].[LOT_RUN] LR
                INNER JOIN [omRESULTS].[dbo].[MEASUREMENT] M ON LR.[RUN_ID] = M.[RUN_ID]
                INNER JOIN [omRESULTS].[dbo].[AROL_MEASUREMENT] AM ON M.[MEASUREMENT_ID] = AM.[MEASUREMENT_ID]
                """
                st.session_state.tab4_total_rows = pd.read_sql_query(count_query, conn).iloc[0]["TotalRows"]
                conn.close()
                st.session_state.tab4_page = 1
            except Exception as e:
                st.error(f"Connection failed or query error: {e}")
            finally:
                progress_bar.empty()

    # 데이터 표시
    if st.session_state.tab4_total_rows > 0:
        rows_per_page = TAB4_ROWS_PER_PAGE
        offset = (st.session_state.tab4_page - 1) * rows_per_page
        data = tab4_get_data(offset, rows_per_page, TAB4_IP_ADDRESS, st.session_state.tab4_port)

        if data is not None:
            # 초 단위 변환
            data['RUN_START_TIME'] = pd.to_datetime(data['RUN_START_TIME'], unit='s', errors='coerce')
            data['RUN_END_TIME'] = pd.to_datetime(data['RUN_END_TIME'], unit='s', errors='coerce')

            # 밀리초 단위 변환 (필요시)
            data['RUN_START_TIME'] = pd.to_datetime(data['RUN_START_TIME'], unit='ms', errors='coerce')
            data['RUN_END_TIME'] = pd.to_datetime(data['RUN_END_TIME'], unit='ms', errors='coerce')


            # Display the formatted dataframe
            st.dataframe(data)

            # CSV Download button
            csv = data.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="Download data as CSV",
                data=csv,
                file_name='tab4_data.csv',
                mime='text/csv',
                key="tab4_download"
            )
        else:
            st.info("No data available to display. Please load the data using the 'Load Data' button.")


            
            
with tab5:
    st.header("Diagnostisc Result Dashboard")

    # 기본 설정값
    TAB5_IP_ADDRESS = ip_address
    TAB5_DATABASE = "DIAGNOSTICS"
    TAB5_USERNAME = "sa"
    TAB5_PASSWORD = "atlas"
    TAB5_ROWS_PER_PAGE = 500

    # 상태 초기화
    if "tab5_total_rows" not in st.session_state:
        st.session_state.tab5_total_rows = 0
    if "tab5_page" not in st.session_state:
        st.session_state.tab5_page = 1
    if "tab5_port" not in st.session_state:
        st.session_state.tab5_port = None

    # TEST_NAME 값을 가져오는 함수
    def get_test_names(ip_address, port):
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={ip_address},{port};"
                f"DATABASE={TAB5_DATABASE};"
                f"UID={TAB5_USERNAME};"
                f"PWD={TAB5_PASSWORD};"
            )
            query = """
            SELECT DISTINCT T.TEST_NAME
            FROM [DIAGNOSTICS].[dbo].[TEST_RUN] T
            INNER JOIN [DIAGNOSTICS].[dbo].[RESULT] R ON R.TEST_RUN_ID = T.TEST_RUN_ID
            """
            test_names = pd.read_sql_query(query, conn)["TEST_NAME"].tolist()
            conn.close()
            return test_names
        except Exception as e:
            st.error(f"Failed to fetch TEST_NAME values: {e}")
            return []

    # 필터링된 데이터를 가져오는 함수
    def get_filtered_data(offset, fetch, ip_address, port, selected_test_name):
        try:
            conn = pyodbc.connect(
                f"DRIVER={{ODBC Driver 17 for SQL Server}};"
                f"SERVER={ip_address},{port};"
                f"DATABASE={TAB5_DATABASE};"
                f"UID={TAB5_USERNAME};"
                f"PWD={TAB5_PASSWORD};"
            )
            start_row = offset + 1
            end_row = offset + fetch
            query = """
            SELECT *
            FROM (
                SELECT
                    ROW_NUMBER() OVER (ORDER BY T.TEST_NAME ASC, S.START_TIME DESC) AS RowNum,
                    T.TEST_NAME,
                    R.MEASURED_VALUE,
                    R.UNITS,
                    R.HIGH_LIMIT,
                    R.LOW_LIMIT,
                    S.START_TIME,
                    S.END_TIME
                FROM [DIAGNOSTICS].[dbo].[RESULT] R
                INNER JOIN [DIAGNOSTICS].[dbo].[TEST_RUN] T ON R.TEST_RUN_ID = T.TEST_RUN_ID
                INNER JOIN [DIAGNOSTICS].[dbo].[SESSION] S ON T.SESSION_ID = S.SESSION_ID
                WHERE T.TEST_NAME = ?
            ) AS Results
            WHERE RowNum BETWEEN ? AND ?
            """
            data = pd.read_sql_query(query, conn, params=[selected_test_name, start_row, end_row])
            conn.close()
            return data
        except Exception as e:
            st.error(f"Connection failed or query error: {e}")
            return None

    # Load 버튼 클릭 시
    if st.button("Load Data", key="tab5_search"):
        progress_bar = st.progress(0)  # 프로그래스 바 초기화
        st.session_state.tab5_port = find_sql_server_port(TAB5_IP_ADDRESS)

        if not st.session_state.tab5_port:
            st.error("Could not find the SQL Server port.")
        else:
            try:
                for i in range(100):  # 프로그래스 바 애니메이션
                    time.sleep(0.02)
                    progress_bar.progress(i + 1)
                st.session_state.tab5_test_names = get_test_names(TAB5_IP_ADDRESS, st.session_state.tab5_port)
                st.session_state.tab5_page = 1
            except Exception as e:
                st.error(f"Connection failed or query error: {e}")
            finally:
                progress_bar.empty()  # 프로그래스 바 제거

    # TEST_NAME 드롭다운
    if "tab5_test_names" in st.session_state and st.session_state.tab5_test_names:
        sorted_test_names = sorted(st.session_state.tab5_test_names)
        selected_test_name = st.selectbox("Select a TEST_NAME to filter:", sorted_test_names, key="tab5_test_name_dropdown")

        if selected_test_name:
            offset = (st.session_state.tab5_page - 1) * TAB5_ROWS_PER_PAGE
            progress_bar = st.progress(0)  # 프로그래스 바 초기화
            try:
                for i in range(100):  # 데이터 로드 중 프로그래스 바 애니메이션
                    time.sleep(0.02)
                    progress_bar.progress(i + 1)

                data = get_filtered_data(offset, TAB5_ROWS_PER_PAGE, TAB5_IP_ADDRESS, st.session_state.tab5_port, selected_test_name)

                if data is not None and not data.empty:
                    st.dataframe(data)

                    # CSV 다운로드 버튼
                    csv = data.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Download data as CSV",
                        data=csv,
                        file_name='tab5_data.csv',
                        mime='text/csv',
                        key="tab5_download"
                    )
                else:
                    st.info("No data available for this page.")
            except Exception as e:
                st.error(f"Error loading data: {e}")
            finally:
                progress_bar.empty()  # 프로그래스 바 제거

            # 페이지 네비게이션
            if st.session_state.tab5_total_rows > 0:
                total_pages = math.ceil(st.session_state.tab5_total_rows / TAB5_ROWS_PER_PAGE)
                cols = st.columns(3)
                if cols[0].button("◀", key="tab5_prev_page") and st.session_state.tab5_page > 1:
                    st.session_state.tab5_page -= 1
                cols[1].write(f"Page {st.session_state.tab5_page} of {total_pages}")
                if cols[2].button("▶", key="tab5_next_page") and st.session_state.tab5_page < total_pages:
                    st.session_state.tab5_page += 1
