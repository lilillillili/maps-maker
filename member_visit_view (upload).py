# member_visit_view_v4.py - 구글 지도 API 버전

import pandas as pd
import json
import requests
import time

class ExcelToGoogleMap:
    def __init__(self, excel_file_path, google_api_key):
        self.excel_file_path = excel_file_path
        self.google_api_key = google_api_key
        self.df = None
        self.company_locations = []

    def load_excel(self):
        """엑셀 파일을 로드합니다."""
        try:
            if self.excel_file_path.endswith('.csv'):
                self.df = pd.read_csv(self.excel_file_path)
            else:
                self.df = pd.read_excel(self.excel_file_path)
            
            print(f"✅ 엑셀 파일 로드 완료: {len(self.df)}개 행")
            required_columns = ['회원사명', '주소']
            if not all(col in self.df.columns for col in required_columns):
                print(f"❌ 필요한 컬럼('회원사명', '주소')이 없습니다.")
                return False
            return True
        except Exception as e:
            print(f"❌ 엑셀 파일 로드 실패: {e}")
            return False

    def test_google_api_connection(self):
        """구글 API 연결을 테스트합니다."""
        print("🔍 구글 지도 API 연결 테스트 중...")
        test_address = "서울특별시 강남구 테헤란로 152"
        
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            'address': test_address,
            'key': self.google_api_key,
            'language': 'ko',
            'region': 'kr'
        }
        
        try:
            print(f"  📍 테스트 주소: {test_address}")
            print(f"  🔑 API Key: {self.google_api_key[:8]}...")
            
            response = requests.get(url, params=params, timeout=10)
            print(f"  📡 응답 코드: {response.status_code}")
            
            if response.status_code == 200:
                result = response.json()
                status = result.get('status')
                
                if status == 'OK':
                    print("  ✅ API 연결 성공!")
                    return True
                elif status == 'REQUEST_DENIED':
                    print(f"  ❌ API 키 오류: {result.get('error_message', 'API 키가 유효하지 않습니다.')}")
                elif status == 'OVER_QUERY_LIMIT':
                    print("  ❌ API 사용량 초과")
                else:
                    print(f"  ❌ API 응답 오류: {status} - {result.get('error_message', '')}")
            else:
                print(f"  ❌ HTTP 오류: {response.status_code}")
                print(f"  📄 응답 내용: {response.text}")
                
        except Exception as e:
            print(f"  💥 연결 테스트 실패: {e}")
            
        return False

    def geocode_address_google(self, address):
        """구글 Geocoding API를 사용해 주소를 좌표로 변환합니다."""
        cleaned_address = address.strip().replace('\n', ' ').replace('\r', ' ')
        
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {
            'address': cleaned_address,
            'key': self.google_api_key,
            'language': 'ko',  # 한국어 우선
            'region': 'kr'     # 한국 지역 우선
        }
        
        try:
            # API 요청 제한을 위한 딜레이 (구글은 초당 50회 제한)
            time.sleep(0.05)
            
            response = requests.get(url, params=params, timeout=15)
            
            if response.status_code != 200:
                print(f"  ❌ HTTP 오류 {response.status_code}: {response.text}")
                return None
                
            result = response.json()
            status = result.get('status')
            
            if status == 'OK' and result.get('results'):
                # 첫 번째 결과 사용 (가장 정확한 결과)
                location = result['results'][0]['geometry']['location']
                return {
                    'lat': float(location['lat']), 
                    'lng': float(location['lng'])
                }
            elif status == 'ZERO_RESULTS':
                print(f"  ⚠️ 검색 결과 없음: '{cleaned_address}'")
                return None
            elif status == 'OVER_QUERY_LIMIT':
                print(f"  ⏱️ API 사용량 초과 - 1초 대기 후 재시도")
                time.sleep(1)
                return self.geocode_address_google(address)  # 재시도
            elif status == 'REQUEST_DENIED':
                print(f"  🚫 API 키 오류: {result.get('error_message', '')}")
                return None
            else:
                print(f"  ⚠️ API 오류 {status}: {result.get('error_message', '')}")
                return None
                
        except requests.exceptions.Timeout:
            print(f"  ⏰ API 요청 타임아웃")
            return None
        except requests.exceptions.RequestException as e:
            print(f"  💥 API 요청 오류: {e}")
            return None
        except Exception as e:
            print(f"  💥 알 수 없는 오류: {e}")
            return None

    def process_addresses(self):
        """모든 주소를 처리하여 좌표로 변환합니다."""
        if self.df is None: 
            return False
            
        # API 연결 테스트 먼저 실행
        if not self.test_google_api_connection():
            print("\n❌ 구글 지도 API 연결 실패. 다음을 확인해주세요:")
            print("   1. https://console.cloud.google.com 에서 프로젝트 생성")
            print("   2. Maps JavaScript API 및 Geocoding API 활성화")
            print("   3. API 키 생성 및 정확한 입력")
            print("   4. 결제 정보 등록 (무료 사용량: 월 $200)")
            return False
        
        print("\n🔄 구글 지도 API로 주소를 좌표로 변환 중...")
        print("=" * 60)
        
        success_count = 0
        fail_count = 0
        
        for index, row in self.df.iterrows():
            company_name = str(row.get('회원사명', '')).strip()
            address = str(row.get('주소', '')).strip()

            if not company_name or not address or address == 'nan':
                continue

            print(f"\n📋 처리 중 ({index+1}/{len(self.df)}): {company_name}")
            print(f"  📍 주소: {address}")

            coords = self.geocode_address_google(address)
            
            if coords:
                self.company_locations.append({
                    'name': company_name,
                    'address': address,
                    'lat': coords['lat'],
                    'lng': coords['lng']
                })
                print(f"  ✅ 성공: ({coords['lat']:.6f}, {coords['lng']:.6f})")
                success_count += 1
            else:
                print(f"  ❌ 실패: 좌표를 찾을 수 없습니다.")
                fail_count += 1
            
        print("\n" + "=" * 60)
        print(f"🎉 처리 완료: 성공 {success_count}개, 실패 {fail_count}개")
        return success_count > 0

    def generate_html(self, output_path="KESSIA_회원사_지도_구글.html"):
        """구글 지도와 테이블이 포함된 HTML 파일을 생성합니다."""
        if not self.company_locations:
            print("❌ 처리된 위치 데이터가 없어 HTML을 생성할 수 없습니다.")
            return

        avg_lat = sum(loc['lat'] for loc in self.company_locations) / len(self.company_locations)
        avg_lng = sum(loc['lng'] for loc in self.company_locations) / len(self.company_locations)
        
        # 회사 위치를 JavaScript 형태로 변환
        locations_js = json.dumps(self.company_locations, ensure_ascii=False)
        
        table_rows_html = ""
        for i, location in enumerate(self.company_locations, 1):
            table_rows_html += f"""
            <tr onclick="panToMarker({i-1})">
                <td>{i}</td>
                <td>{location['name']}</td>
                <td>{location['address']}</td>
            </tr>
            """
        
        html_content = f'''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>KESSIA 회원사 위치 지도 (Google Maps Ver.)</title>
    <style>
        body {{ 
            font-family: 'Malgun Gothic', Arial, sans-serif; 
            margin: 0; 
            background-color: #f5f5f5; 
        }}
        
        .header {{ 
            text-align: center; 
            padding: 20px; 
            background: linear-gradient(135deg, #4285F4, #34A853);
            color: white; 
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        
        .main-content {{ 
            display: flex; 
            gap: 20px; 
            padding: 20px; 
            max-width: 1400px;
            margin: 0 auto;
        }}
        
        .map-container {{ 
            flex: 7; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.1); 
            border-radius: 12px; 
            overflow: hidden; 
            background: white;
        }}
        
        #map {{ 
            height: 85vh; 
            width: 100%;
        }}
        
        .list-container {{ 
            flex: 3; 
            background: white; 
            box-shadow: 0 4px 20px rgba(0,0,0,0.1); 
            border-radius: 12px; 
            padding: 25px; 
            height: 85vh; 
            display: flex; 
            flex-direction: column; 
        }}
        
        .stats-info {{ 
            background: linear-gradient(135deg, #e3f2fd, #f3e5f5);
            border-left: 4px solid #4285F4;
            padding: 15px; 
            margin-bottom: 20px; 
            border-radius: 8px; 
            font-size: 14px;
        }}
        
        .table-wrapper {{ 
            overflow-y: auto; 
            flex-grow: 1; 
            border-radius: 8px;
            border: 1px solid #e0e0e0;
        }}
        
        table {{ 
            width: 100%; 
            border-collapse: collapse; 
        }}
        
        th, td {{ 
            padding: 12px 10px; 
            text-align: left; 
            border-bottom: 1px solid #eee; 
        }}
        
        th {{ 
            background: linear-gradient(135deg, #f8f9fa, #e9ecef);
            position: sticky; 
            top: 0; 
            font-weight: 600;
            z-index: 10;
        }}
        
        tbody tr:hover {{ 
            background-color: #e3f2fd; 
            cursor: pointer; 
            transform: translateX(2px);
            transition: all 0.2s ease;
        }}
        
        .loading {{ 
            text-align: center; 
            padding: 50px; 
            color: #666; 
        }}

        @media (max-width: 768px) {{
            .main-content {{ 
                flex-direction: column; 
                padding: 10px;
            }}
            .map-container, .list-container {{ 
                flex: none; 
                height: 50vh; 
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>🏢 KESSIA 회원사 위치 지도</h1>
        <p style="margin: 5px 0; opacity: 0.9;">Google Maps API 기반 정확한 위치 정보</p>
        <p style="margin: 0; font-size: 14px; opacity: 0.8;">총 {len(self.company_locations)}개 회원사</p>
    </div>
    
    <div class="main-content">
        <div class="map-container">
            <div id="map">
                <div class="loading">🔄 구글 지도 로딩 중...</div>
            </div>
        </div>
        
        <div class="list-container">
            <h2 style="margin-top: 0; color: #333;">📋 회원사 목록</h2>
            
            <div class="stats-info">
                <strong>🎯 지도 정보</strong><br>
                • 총 회원사: {len(self.company_locations)}개<br>
                • API: Google Maps<br>
                • 클릭하면 해당 위치로 이동합니다
            </div>
            
            <div class="table-wrapper">
                <table>
                    <thead>
                        <tr>
                            <th style="width: 50px;">순번</th>
                            <th style="width: 35%;">회원사명</th>
                            <th>주소</th>
                        </tr>
                    </thead>
                    <tbody>{table_rows_html}</tbody>
                </table>
            </div>
        </div>
    </div>

    <script>
        let map;
        let markers = [];
        const companyLocations = {locations_js};

        function initMap() {{
            // 지도 초기화
            const centerLat = {avg_lat};
            const centerLng = {avg_lng};
            
            map = new google.maps.Map(document.getElementById('map'), {{
                zoom: 8,
                center: {{ lat: centerLat, lng: centerLng }},
                mapTypeId: 'roadmap',
                styles: [
                    {{
                        featureType: 'poi',
                        elementType: 'labels',
                        stylers: [{{ visibility: 'on' }}]
                    }}
                ]
            }});

            // 마커 추가
            const bounds = new google.maps.LatLngBounds();
            
            companyLocations.forEach((location, index) => {{
                const marker = new google.maps.Marker({{
                    position: {{ lat: location.lat, lng: location.lng }},
                    map: map,
                    title: location.name,
                    animation: google.maps.Animation.DROP
                }});

                // 정보창 생성
                const infoWindow = new google.maps.InfoWindow({{
                    content: `
                        <div style="padding: 10px; max-width: 300px;">
                            <h3 style="margin: 0 0 8px 0; color: #333; font-size: 16px;">
                                🏢 ${{location.name}}
                            </h3>
                            <p style="margin: 0; color: #666; font-size: 14px; line-height: 1.4;">
                                📍 ${{location.address}}
                            </p>
                        </div>
                    `
                }});

                // 마커 클릭 이벤트
                marker.addListener('click', () => {{
                    // 다른 정보창들 닫기
                    markers.forEach(m => m.infoWindow.close());
                    // 현재 정보창 열기
                    infoWindow.open(map, marker);
                }});

                // 마커와 정보창 저장
                markers.push({{ marker: marker, infoWindow: infoWindow }});
                
                // 경계 확장
                bounds.extend(marker.getPosition());
            }});

            // 모든 마커가 보이도록 지도 조정
            if (companyLocations.length > 0) {{
                map.fitBounds(bounds);
                
                // 최대 줌 레벨 제한 (너무 가까이 가지 않도록)
                const listener = google.maps.event.addListener(map, 'idle', () => {{
                    if (map.getZoom() > 16) map.setZoom(16);
                    google.maps.event.removeListener(listener);
                }});
            }}
        }}

        // 테이블에서 클릭했을 때 해당 마커로 이동
        function panToMarker(index) {{
            if (index >= 0 && index < companyLocations.length) {{
                const location = companyLocations[index];
                
                // 지도 이동 및 줌
                map.panTo({{ lat: location.lat, lng: location.lng }});
                map.setZoom(16);
                
                // 해당 마커의 정보창 열기
                setTimeout(() => {{
                    markers.forEach(m => m.infoWindow.close()); // 다른 정보창 닫기
                    markers[index].infoWindow.open(map, markers[index].marker);
                }}, 500);
            }}
        }}

        // 구글 지도 로드 실패 시 처리
        window.gm_authFailure = function() {{
            document.getElementById('map').innerHTML = 
                '<div style="padding: 50px; text-align: center; color: #d32f2f;">' +
                '<h3>❌ 구글 지도 로드 실패</h3>' +
                '<p>API 키를 확인해주세요.</p>' +
                '</div>';
        }}
    </script>
    
    <!-- 구글 지도 API 로드 (여기에 실제 API 키를 입력하세요) -->
    <script async defer 
        src="https://maps.googleapis.com/maps/api/js?key={self.google_api_key}&callback=initMap&language=ko&region=KR">
    </script>
</body>
</html>'''
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"\n✅ 구글 지도 기반 HTML 생성 완료: {output_path}")
        print(f"🌐 브라우저에서 파일을 열어 지도를 확인하세요!")

    def run(self):
        """전체 프로세스를 실행합니다."""
        print("🚀 KESSIA 회원사 지도 생성 프로그램 (Google Maps Ver.)")
        print("=" * 60)
        
        if not self.load_excel():
            return
        if not self.process_addresses():
            return
        self.generate_html()
        
        print("\n" + "=" * 60)
        print("🎉 모든 작업이 완료되었습니다!")
        print("💡 Tip: 생성된 HTML 파일을 브라우저에서 열어보세요.")

# --- 실행 부분 ---
if __name__ == "__main__":
    # 1. 엑셀/CSV 파일 경로를 지정하세요.
    excel_file = "방문 회원사 목록 DB_20250805.xlsx"
    
    # 2. 여기에 발급받은 구글 API 키를 입력하세요.
    google_api_key = "YOUR_GOOGLE_API_KEY"
    
    if not google_api_key or google_api_key == "YOUR_GOOGLE_API_KEY":
        print("🛑 [안내] 구글 API 키를 설정해주세요!")
        print("")
        print("📝 구글 지도 API 키 발급 방법:")
        print("1. https://console.cloud.google.com 접속")
        print("2. 새 프로젝트 생성 또는 기존 프로젝트 선택")
        print("3. 'API 및 서비스' > 'API 라이브러리' 이동")
        print("4. 'Maps JavaScript API'와 'Geocoding API' 활성화")
        print("5. 'API 및 서비스' > '사용자 인증 정보'에서 API 키 생성")
        print("6. 생성된 API 키를 위의 google_api_key 변수에 입력")
        print("")
        print("💰 비용: 월 $200 무료 크레딧 (약 28,500회 무료)")
        print("🔒 보안: API 키 제한 설정 권장")
    else:
        mapper = ExcelToGoogleMap(excel_file, google_api_key)
        mapper.run()