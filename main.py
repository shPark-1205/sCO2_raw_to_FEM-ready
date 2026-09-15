import pandas as pd
import numpy as np
import os
import re
from datetime import datetime

# --- [사용자 설정] ---
CONFIG = {
    # 1. 데이터 최상위 폴더 경로
    'base_dir': r'E:\HTL\03.PersonalResearch\Supercritical_CO2\04.Raw data\03. Supercritical\20260914_7.771MPa_31.3C_40000',
}


def extract_date_from_str(text):
    match = re.search(r'(20\d{2})(\d{2})(\d{2})', text)
    if match:
        return datetime.strptime(match.group(0), "%Y%m%d").date()
    return datetime.now().date()


def parse_time_with_date_injection(time_val, base_date):
    if pd.isna(time_val): return pd.NaT
    if isinstance(time_val, datetime):
        if time_val.year == 1900:
            return datetime.combine(base_date, time_val.time())
        return time_val
    if hasattr(time_val, 'hour'):
        return datetime.combine(base_date, time_val)
    time_str = str(time_val).strip()
    try:
        return datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
    except:
        pass
    try:
        return datetime.combine(base_date, datetime.strptime(time_str, "%H:%M:%S").time())
    except:
        return pd.NaT


def process_rtd_dynamic(folders, base_date):
    """
    RTD 데이터 동적 처리:
    1. 파일의 열 개수를 통해 센서 수(N) 계산 ( (전체열-1) / 3 )
    2. 기본 10열(T, t, R 각각)을 맞추기 위해 Gap = 10 - N 계산
    3. Time + [T(N)+Gap] + [t(N)+Gap] + [R(N)+Gap] 구조 생성
    4. A2, A3 시간 읽어서 TimeInfo 저장
    """
    all_processed_data = []
    global_start = None
    global_end = None

    folders.sort()

    for folder in folders:
        files = sorted([f for f in os.listdir(folder) if f.endswith('.xlsx') and not f.startswith('~$')])

        for file in files:
            path = os.path.join(folder, file)
            try:
                # header=0 (1행 헤더)
                df_source = pd.read_excel(path, header=0)
                if len(df_source) < 2: continue

                # --- [동적 센서 수 감지] ---
                # A열(Time) 제외한 나머지 열 개수
                total_data_cols = df_source.shape[1] - 1
                num_sensors = total_data_cols // 3

                # 기본 슬롯 10개 기준, 부족한 만큼 공백 추가
                # 만약 센서가 10개 넘으면 공백 없음(0)
                gap_count = max(0, 10 - num_sensors)

                # 시간 정보 읽기
                start_val = df_source.iloc[0, 0]
                end_val = df_source.iloc[1, 0]

                # Global 시간 갱신
                s_dt = parse_time_with_date_injection(start_val, base_date)
                e_dt = parse_time_with_date_injection(end_val, base_date)
                if global_start is None or (s_dt and s_dt < global_start): global_start = s_dt
                if global_end is None or (e_dt and e_dt > global_end): global_end = e_dt

                # 데이터 슬라이싱 (31~57행 -> iloc 29:56)
                slice_end = min(56, len(df_source))
                if slice_end <= 29: continue
                df_sliced = df_source.iloc[29:slice_end].reset_index(drop=True)

                # --- 데이터 재배치 ---

                # 1) TimeInfo 열 (dtype=object)
                time_col = pd.Series([np.nan] * len(df_sliced), dtype=object)
                time_col[0] = start_val
                if len(time_col) > 1: time_col[1] = end_val

                # 2) 데이터 분리 (T, t, R)
                # 구조: [Time] [T...T] [t...t] [R...R]
                # 인덱스: 0     1~N    1+N~2N  1+2N~3N
                idx_t_start = 1
                idx_time_start = 1 + num_sensors
                idx_r_start = 1 + 2 * num_sensors

                t_data = df_sliced.iloc[:, idx_t_start: idx_t_start + num_sensors]
                time_data = df_sliced.iloc[:, idx_time_start: idx_time_start + num_sensors]
                r_data = df_sliced.iloc[:, idx_r_start: idx_r_start + num_sensors]

                # 3) 빈 열(Gap) 생성
                if gap_count > 0:
                    gap = pd.DataFrame(np.nan, index=df_sliced.index, columns=[f'Gap{i}' for i in range(gap_count)])
                else:
                    gap = pd.DataFrame()  # 빈 DF

                # 4) 가로 병합: Time + (T+Gap) + (t+Gap) + (R+Gap)
                # 요청사항: "R 데이터 뒤에도 똑같이 공백 열이 있어야 해" 반영
                df_combined = pd.concat([time_col, t_data, gap, time_data, gap, r_data, gap], axis=1)

                all_processed_data.append(df_combined)

                # 빈 행 추가
                blank_row = pd.DataFrame([[np.nan] * df_combined.shape[1]], columns=df_combined.columns)
                all_processed_data.append(blank_row)

            except Exception as e:
                print(f"    [Error] {file}: {e}")

    if not all_processed_data:
        return None, None, None

    if len(all_processed_data) > 0: all_processed_data.pop()

    final_df = pd.concat(all_processed_data, ignore_index=True)
    return final_df, global_start, global_end


def process_filter_data(path, base_date, start_dt, end_dt):
    try:
        df = pd.read_excel(path) if path.endswith('.xlsx') else pd.read_csv(path)
    except:
        return None

    time_col = next((c for c in df.columns if 'time' in c.lower()), None)
    if not time_col: return None

    temp_dates = df[time_col].apply(lambda x: parse_time_with_date_injection(x, base_date))

    if start_dt and end_dt:
        mask = (temp_dates >= start_dt) & (temp_dates <= end_dt)
        filtered = df.loc[mask].copy()
    else:
        filtered = pd.DataFrame()

    return filtered


def main():
    base_dir = CONFIG['base_dir']
    rtd_dir = os.path.join(base_dir, 'RTD')
    hx_dir = os.path.join(base_dir, 'HX')

    sys_files = [f for f in os.listdir(base_dir) if f.startswith('_SYS') and f.endswith('.xlsx')]
    sys_path = os.path.join(base_dir, sys_files[0]) if sys_files else None

    rtd_groups = {}
    if os.path.exists(rtd_dir):
        for d in os.listdir(rtd_dir):
            if '-r' in d and d.endswith('mm'):
                match = re.search(r'_RTD-(.+)-r\d+mm', d)
                if match:
                    key = match.group(1)
                    if key not in rtd_groups: rtd_groups[key] = []
                    rtd_groups[key].append(os.path.join(rtd_dir, d))

    print(f"--- 총 {len(rtd_groups)}개의 실험 세트 처리 시작 ---\n")

    for group_key, folders in rtd_groups.items():
        print(f">>> 처리 중: {group_key}")

        sample_file = os.listdir(folders[0])[0]
        base_date = extract_date_from_str(sample_file)

        # A. RTD 처리 (동적 Gap 계산)
        rtd_df, start_dt, end_dt = process_rtd_dynamic(folders, base_date)
        if rtd_df is None: continue

        print(f"    시간 범위: {start_dt.time()} ~ {end_dt.time()}")

        # B. System 처리
        sys_df = process_filter_data(sys_path, base_date, start_dt, end_dt)

        # C. Heater 처리
        heater_df = None
        if os.path.exists(hx_dir):
            hx_candidates = [f for f in os.listdir(hx_dir) if group_key in f and f.endswith('.xlsx')]
            if hx_candidates:
                heater_df = process_filter_data(os.path.join(hx_dir, hx_candidates[0]), base_date, start_dt, end_dt)

        # D. 저장
        output_filename = f"{base_date.strftime('%Y%m%d')}_{group_key}_Merged.xlsx"
        output_path = os.path.join(base_dir, output_filename)

        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 1. RTD 저장 (헤더 없음)
                rtd_df.to_excel(writer, sheet_name='MergedData', startrow=0, startcol=0, index=False, header=False)

                # 2. System 저장 (RTD 끝난 후 1열 띄움)
                # RTD 폭 = 1(Time) + 10(T) + 10(t) + 10(R) = 31 (만약 센서가 10 이하면 항상 31)
                # rtd_df.shape[1]을 쓰면 계산된 폭이 나옴
                sys_start_col = rtd_df.shape[1] + 1
                if sys_df is not None and not sys_df.empty:
                    sys_df.to_excel(writer, sheet_name='MergedData', startrow=0, startcol=sys_start_col, index=False,
                                    header=False)

                # 3. Heater 저장 (System 끝난 후 1열 띄움)
                sys_width = len(sys_df.columns) if (sys_df is not None and not sys_df.empty) else 0
                heater_start_col = sys_start_col + sys_width + 1
                if heater_df is not None and not heater_df.empty:
                    heater_df.to_excel(writer, sheet_name='MergedData', startrow=0, startcol=heater_start_col,
                                       index=False, header=False)

            print(f"    [생성 완료] {output_filename}")

        except Exception as e:
            print(f"    [저장 실패] {e}")
        print("")


if __name__ == "__main__":
    main()