import pandas as pd
import folium
from geopy.geocoders import Nominatim
import time

def analyze_2025_trips():
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel('./출장보고서2005.xlsx')
        print("파일 읽기 성공!")

        # 기본 데이터 정보 출력
        print("\n=== 데이터 기본 정보 ===")
        print("전체 출장 건수:", len(df))
        print("\n컬럼 목록:")
        print(df.columns.tolist())
        
        # 출장인원 통계
        print("\n=== 출장인원 통계 ===")
        if '출장인원(이름)' in df.columns:
            all_people = []
            for names in df['출장인원(이름)'].dropna():
                all_people.extend([name.strip() for name in str(names).split('/')])
            unique_people = sorted(list(set(all_people)))
            print("전체 출장인원 목록:")
            for person in unique_people:
                print(f"- {person}")

        # 방문 지역 통계
        print("\n=== 방문 지역 통계 ===")
        visit_counts = df['출장장소(시군구)'].value_counts()
        print("지역별 방문 횟수:")
        print(visit_counts)

        # 방문 횟수별 지역 분류
        visit_categories = {
            '4회 이상 방문': [],
            '3회 방문': [],
            '2회 방문': [],
            '1회 방문': []
        }

        for location, count in visit_counts.items():
            if count >= 4:
                visit_categories['4회 이상 방문'].append(location)
            elif count == 3:
                visit_categories['3회 방문'].append(location)
            elif count == 2:
                visit_categories['2회 방문'].append(location)
            else:
                visit_categories['1회 방문'].append(location)

        # 지도 생성
        m = folium.Map(location=[36.5, 127.5], zoom_start=7)

        # 범례 표 HTML 생성
        legend_html = '''
        <div style="position: fixed; 
                    top: 50px; 
                    right: 50px; 
                    width: 200px;
                    height: auto;
                    background-color: white;
                    border: 2px solid grey;
                    z-index: 1000;
                    padding: 10px;
                    font-family: Arial;">
            <h4 style="text-align: center; margin: 0 0 10px 0;">방문 횟수별 지역</h4>
            <table style="width: 100%; border-collapse: collapse;">
        '''
        
        # 색상 매핑
        colors = {
            '4회 이상 방문': 'black',
            '3회 방문': 'green',
            '2회 방문': 'red',
            '1회 방문': 'blue'
        }

        for category, locations in visit_categories.items():
            if locations:  # 해당 카테고리에 지역이 있는 경우만 표시
                legend_html += f'''
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 5px;">
                        <div style="width: 15px; 
                                  height: 15px; 
                                  background-color: {colors[category]}; 
                                  display: inline-block;
                                  margin-right: 5px;"></div>
                        <b>{category}</b>
                    </td>
                </tr>
                <tr style="border-bottom: 1px solid #ddd;">
                    <td style="padding: 5px 5px 10px 20px; font-size: 0.9em;">
                        {", ".join(sorted(locations))}
                    </td>
                </tr>
                '''

        legend_html += '''
            </table>
        </div>
        '''

        # 제목과 범례를 지도에 추가
        title_html = '''
        <h3 align="center" style="font-size:20px">
            <b>2025년 경매사 방문지역</b>
        </h3>
        '''
        
        m.get_root().html.add_child(folium.Element(title_html + legend_html))

        geolocator = Nominatim(user_agent="my_agent")

        # 주소 검색을 위한 보조 정보
        address_corrections = {
            '고창군': '전라북도 고창군',
            '부안군': '전라북도 부안군',
            '진주시': '경상남도 진주시',
            '보성군': '전라남도 보성군',
            '나주시': '전라남도 나주시',
            '천안시': '충청남도 천안시',
            '예산군': '충청남도 예산군',
            '홍성군': '충청남도 홍성군',
            '영주시': '경상북도 영주시',
            '서산시': '충청남도 서산시',
            '당진시': '충청남도 당진시',
            '영천시': '경상북도 영천시',
            '문경시': '경상북도 문경시',
            '상주시': '경상북도 상주시',
            '김천시': '경상북도 김천시'
        }

        # 팝업 정보 생성 부분 수정
        detailed_items = {
            '대구시': '참외 작목현황 조사 및 재배농가 현장점검',
            '밀양': '하우스감자 시세 동향 조사 및 출하 상담, 재배현장 점검',
            '해남군': '봉동배추, 대파 출하독려 및 작황조사',
            '김천시': '사과 출하 독려 및 저장현황 점검',
            '논산시': '배 출하 독려 및 저장량 조사',
            '남원시': '사과 출하 독려 및 시장동향 파악',
            '예산군': '사과 출하 독려 및 저장현황 조사',
            '영천시': '깐마늘 출하 독려 및 작업현장 점검',
            '사천시': '통연근 출하 독려 및 작황조사',
            '예산군': '쪽파 출하독려 및 재배현황 점검',
            '금산군': '상추, 깻잎 출하 독려 및 작황점검',
            '옥천군': '상추, 깻잎 출하 독려 및 시장동향 파악',
            '고창군': '쪽파 출하독려 및 재배현황 조사'
        }

        # 각 위치에 마커 추가
        failed_locations = []
        for location, count in visit_counts.items():
            try:
                search_location = address_corrections.get(location, f"대한민국 {location}")
                loc = geolocator.geocode(search_location)
                time.sleep(1)
                
                if loc:
                    # 방문 횟수에 따른 아이콘 색상
                    if count >= 4:
                        icon_color = 'black'
                    elif count == 3:
                        icon_color = 'green'
                    elif count == 2:
                        icon_color = 'red'
                    else:
                        icon_color = 'blue'

                    # 해당 위치의 방문 정보 수집
                    location_visits = df[df['출장장소(시군구)'] == location]
                    
                    # 팝업 텍스트 생성
                    popup_text = f'''
                    <div style="font-size: 16px;">  <!-- 기본 글씨 크기 증가 -->
                        <h3 style="font-size: 22px; margin: 8px 0; color: #2c3e50;"><b>{location}</b></h3>
                        <p style="margin: 8px 0; font-size: 18px;"><b>방문 횟수: {count}회</b></p>
                        <br>
                        <p style="margin: 8px 0; font-size: 18px;"><b>방문 내역:</b></p>
                    '''

                    for _, row in location_visits.iterrows():
                        visit_date = pd.to_datetime(row['출장일자(시작)']).strftime('%Y-%m-%d')
                        
                        # 주요품목 상세 설명
                        detailed_item = detailed_items.get(location, row['주요품목'])
                        popup_text += f'''
                        <div style="margin: 12px 0; background-color: #f8f9fa; padding: 10px; border-radius: 5px;">
                            <p style="margin: 4px 0; font-size: 18px;"><b>{visit_date}</b></p>
                            <p style="margin: 4px 0; font-size: 16px;">• {detailed_item}</p>
                            <p style="margin: 4px 0; font-size: 16px;">• 출장인원: {row['출장인원(이름)']}</p>
                        </div>
                        '''

                    popup_text += '</div>'

                    # 마커 추가
                    folium.Marker(
                        location=[loc.latitude, loc.longitude],
                        popup=folium.Popup(popup_text, max_width=450),  # 팝업 크기 증가
                        tooltip=f"{location} ({count}회 방문)",
                        icon=folium.Icon(color=icon_color, icon='info-sign')
                    ).add_to(m)
                else:
                    failed_locations.append(location)
            except Exception as e:
                print(f"{location} 처리 중 오류: {str(e)}")
                failed_locations.append(location)

        # 지도 저장
        m.save('./business_trips_map_2025.html')
        print("\n2025년도 방문 지도가 'business_trips_map_2025.html' 파일로 저장되었습니다.")

        # 실패한 위치 정보 출력
        if failed_locations:
            print("\n위치를 찾지 못한 지역:")
            for loc in failed_locations:
                print(f"- {loc}")

        # 엑셀 파일로 정리된 데이터 저장
        df.to_excel('./business_trips_2025_analyzed.xlsx', index=False)
        print("\n분석된 데이터가 'business_trips_2025_analyzed.xlsx' 파일로 저장되었습니다.")

    except Exception as e:
        print(f"처리 중 오류 발생: {str(e)}")

if __name__ == "__main__":
    analyze_2025_trips() 