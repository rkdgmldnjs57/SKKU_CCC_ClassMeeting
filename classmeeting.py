import pandas as pd
import re

def extract_info(column):
    # 정규표현식을 사용하여 문자열 분리
    pattern = r'(\S+)\【(\d+)\】,(\S+)\【(\d+)\】'
    matches = re.findall(pattern, column)
    # 열 분리 후, 결과를 딕셔너리로 반환
    return [{'시간1': match[0], '강의실1': int(match[1]), '시간2': match[2], '강의실2': int(match[3])} for match in matches]

def shift_column_up(df, column_name):
    # 열을 Series로 추출
    target_series = df[column_name]

    # Series를 하나씩 위로 이동
    for i in range(len(target_series) - 1):
        target_series[i] = target_series[i + 1]

    # DataFrame에 업데이트
    df[column_name] = target_series

    return df

def process_excel(file_path):
    # 엑셀 파일 읽기
    df = pd.read_excel(file_path)

    df.drop(df.columns[[0, 2, 3, 4, 6, 8]], axis=1, inplace=True)
    df.drop(df.index[0:3], axis=0, inplace=True)
    df.reset_index(drop = True, inplace=True)
    print(df.head())
    df = shift_column_up(df, df.columns[1])
    df.dropna(axis=0, inplace=True)
    df.reset_index(drop = True, inplace=True)
    print(df)

    # 두 번째 열의 정보 추출
    info = df.iloc[:, 1].apply(extract_info)

    # 새로운 DataFrame 생성
    new_df = pd.DataFrame(info.explode().tolist())

    # 기존 DataFrame과 새로운 DataFrame을 합침
    df = pd.concat([df, new_df], axis=1)
    df.drop(df.columns[1], axis=1, inplace=True)
    df.sort_values([df.columns[2], df.columns[1]], inplace=True)

    # 엑셀 파일로 저장
    df.to_excel('processed_' + file_path, index=False)

    print("두 번째 열의 정보가 성공적으로 추출되어 엑셀 파일에 추가되었습니다.")

# 처리할 엑셀 파일 경로 입력
excel_file_path = '영발원본.xlsx'
process_excel(excel_file_path)
