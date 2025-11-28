import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
import re
from urllib.parse import urljoin, urlparse
import warnings
warnings.filterwarnings('ignore')

class CompanyInfoCollector:
    def __init__(self, excel_file_path):
        self.excel_file_path = excel_file_path
        self.df = None
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
    
    def load_excel(self):
        """엑셀 파일 로드"""
        try:
            self.df = pd.read_excel(self.excel_file_path)
            print(f"엑셀 파일 로드 완료: {len(self.df)}개 행")
            return True
        except Exception as e:
            print(f"엑셀 파일 로드 실패: {e}")
            return False
    
    def search_company_info(self, company_name):
        """회사 정보 검색"""
        try:
            # 네이버 검색 사용
            search_url = f"https://search.naver.com/search.naver?query={company_name} 회사 주소 홈페이지"
            
            response = requests.get(search_url, headers=self.headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            address = self.extract_address(soup, company_name)
            homepage = self.extract_homepage(soup, company_name)
            
            return address, homepage
            
        except Exception as e:
            print(f"{company_name} 검색 중 오류: {e}")
            return "", ""
    
    def extract_address(self, soup, company_name):
        """주소 추출"""
        address_patterns = [
            # 지식패널이나 기업 정보에서 주소 찾기
            '.txt_addr',
            '.addr',
            '.address',
            # 일반적인 주소 패턴
            lambda tag: tag.name and '주소' in tag.get_text() and tag.find_next(),
        ]
        
        # 다양한 선택자로 주소 찾기
        selectors = [
            '.business_info .addr',
            '.company_info .address',
            '.info_group .addr',
            '.detail_info .addr'
        ]
        
        for selector in selectors:
            elements = soup.select(selector)
            for element in elements:
                text = element.get_text().strip()
                if self.is_valid_address(text):
                    return text
        
        # 텍스트에서 주소 패턴 찾기
        address_regex = re.compile(r'[가-힣]+[시도]\s+[가-힣]+[시군구]\s+[가-힣\d\-\s,]+', re.MULTILINE)
        text = soup.get_text()
        matches = address_regex.findall(text)
        
        for match in matches:
            if self.is_valid_address(match):
                return match.strip()
        
        return ""
    
    def extract_homepage(self, soup, company_name):
        """홈페이지 추출"""
        # 링크에서 홈페이지 찾기
        links = soup.find_all('a', href=True)
        
        potential_homepages = []
        
        for link in links:
            href = link['href']
            text = link.get_text().strip()
            
            # 공식 홈페이지로 보이는 링크 찾기
            if any(keyword in text.lower() for keyword in ['홈페이지', 'homepage', '공식', 'www', 'http']):
                if self.is_valid_homepage(href):
                    potential_homepages.append(href)
            
            # URL 자체가 회사와 관련되어 보이는 경우
            if self.is_company_website(href, company_name):
                potential_homepages.append(href)
        
        # 가장 적합한 홈페이지 선택
        if potential_homepages:
            homepage = self.select_best_homepage(potential_homepages, company_name)
            
            # 구인구직 사이트면 실제 홈페이지 크롤링
            if self.is_job_site_url(homepage):
                print(f"  -> 구인사이트에서 실제 홈페이지 찾는 중...")
                actual_homepage = self.extract_homepage_from_job_site(homepage)
                return actual_homepage if actual_homepage else homepage
            
            return homepage
        
        return ""
    
    def is_job_site_url(self, url):
        """구인구직 사이트 URL인지 확인"""
        job_sites = [
            'saramin.co.kr', 'jobkorea.co.kr', 'work.go.kr', 'incruit.com',
            'wanted.co.kr', 'indeed.com', 'linkedin.com', 'jobplanet.co.kr',
            'catch.co.kr', 'alba.co.kr', 'albamon.com'
        ]
        
        url_lower = url.lower()
        return any(job_site in url_lower for job_site in job_sites)
    
    def extract_homepage_from_job_site(self, job_site_url):
        """구인구직 사이트에서 실제 회사 홈페이지 추출"""
        try:
            response = requests.get(job_site_url, headers=self.headers, timeout=10)
            response.raise_for_status()
            
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 사람인에서 홈페이지 추출
            if 'saramin' in job_site_url:
                return self.extract_from_saramin(soup)
            
            # 잡코리아에서 홈페이지 추출
            elif 'jobkorea' in job_site_url:
                return self.extract_from_jobkorea(soup)
            
            # 워크넷에서 홈페이지 추출
            elif 'work.go.kr' in job_site_url:
                return self.extract_from_worknet(soup)
            
            # 인크루트에서 홈페이지 추출
            elif 'incruit' in job_site_url:
                return self.extract_from_incruit(soup)
            
            # 기타 사이트는 일반적인 방법으로
            else:
                return self.extract_homepage_generic(soup)
        
        except Exception as e:
            print(f"  -> 구인사이트 크롤링 오류: {e}")
            return ""
    
    def extract_from_saramin(self, soup):
        """사람인에서 홈페이지 추출"""
        try:
            # 사람인의 회사 정보 섹션에서 홈페이지 찾기
            selectors = [
                '.company_info_list a[href*="http"]',
                '.company_summary a[href*="http"]',
                '.basic_info a[href*="http"]',
                'a[title*="홈페이지"]',
                'a[title*="homepage"]'
            ]
            
            for selector in selectors:
                elements = soup.select(selector)
                for element in elements:
                    href = element.get('href', '')
                    if self.is_company_homepage(href):
                        return href
            
            # 텍스트에서 URL 패턴 찾기
            text = soup.get_text()
            urls = re.findall(r'https?://[^\s<>"]+', text)
            for url in urls:
                if self.is_company_homepage(url):
                    return url
        
        except Exception:
            pass
        
        return ""
    
    def extract_from_jobkorea(self, soup):
        """잡코리아에서 홈페이지 추출"""
        try:
            # 잡코리아의 회사 정보 섹션에서 홈페이지 찾기
            selectors = [
                '.tbList a[href*="http"]',
                '.corpDetail a[href*="http"]',
                '.coInfo a[href*="http"]',
                'a[title*="홈페이지"]'
            ]
            
            for selector in selectors:
                elements = soup.select(selector)
                for element in elements:
                    href = element.get('href', '')
                    if self.is_company_homepage(href):
                        return href
            
            # 텍스트에서 URL 패턴 찾기
            text = soup.get_text()
            urls = re.findall(r'https?://[^\s<>"]+', text)
            for url in urls:
                if self.is_company_homepage(url):
                    return url
        
        except Exception:
            pass
        
        return ""
    
    def extract_from_worknet(self, soup):
        """워크넷에서 홈페이지 추출"""
        try:
            # 워크넷의 회사 정보 섹션에서 홈페이지 찾기
            selectors = [
                '.company_info a[href*="http"]',
                '.detail_info a[href*="http"]',
                'a[title*="홈페이지"]'
            ]
            
            for selector in selectors:
                elements = soup.select(selector)
                for element in elements:
                    href = element.get('href', '')
                    if self.is_company_homepage(href):
                        return href
            
            # 텍스트에서 URL 패턴 찾기
            text = soup.get_text()
            urls = re.findall(r'https?://[^\s<>"]+', text)
            for url in urls:
                if self.is_company_homepage(url):
                    return url
        
        except Exception:
            pass
        
        return ""
    
    def extract_from_incruit(self, soup):
        """인크루트에서 홈페이지 추출"""
        try:
            # 인크루트의 회사 정보 섹션에서 홈페이지 찾기
            selectors = [
                '.company_info a[href*="http"]',
                '.info_box a[href*="http"]',
                'a[title*="홈페이지"]'
            ]
            
            for selector in selectors:
                elements = soup.select(selector)
                for element in elements:
                    href = element.get('href', '')
                    if self.is_company_homepage(href):
                        return href
            
            # 텍스트에서 URL 패턴 찾기
            text = soup.get_text()
            urls = re.findall(r'https?://[^\s<>"]+', text)
            for url in urls:
                if self.is_company_homepage(url):
                    return url
        
        except Exception:
            pass
        
        return ""
    
    def extract_homepage_generic(self, soup):
        """일반적인 방법으로 홈페이지 추출"""
        try:
            # 홈페이지 관련 링크 찾기
            links = soup.find_all('a', href=True)
            
            for link in links:
                href = link.get('href', '')
                text = link.get_text().strip().lower()
                title = link.get('title', '').lower()
                
                if ('홈페이지' in text or 'homepage' in text or 
                    '홈페이지' in title or 'homepage' in title):
                    if self.is_company_homepage(href):
                        return href
            
            # 텍스트에서 URL 패턴 찾기
            text = soup.get_text()
            urls = re.findall(r'https?://[^\s<>"]+', text)
            for url in urls:
                if self.is_company_homepage(url):
                    return url
        
        except Exception:
            pass
        
        return ""
    
    def is_company_homepage(self, url):
        """실제 회사 홈페이지인지 확인"""
        if not url or not url.startswith(('http://', 'https://')):
            return False
        
        url_lower = url.lower()
        
        # 구인사이트는 제외
        job_sites = [
            'saramin.co.kr', 'jobkorea.co.kr', 'work.go.kr', 'incruit.com',
            'wanted.co.kr', 'indeed.com', 'linkedin.com', 'jobplanet.co.kr'
        ]
        
        if any(job_site in url_lower for job_site in job_sites):
            return False
        
        # 기타 제외할 사이트들
        excluded_domains = [
            'naver.com', 'google.com', 'daum.net', 'youtube.com',
            'facebook.com', 'instagram.com', 'twitter.com', 'blog'
        ]
        
        if any(domain in url_lower for domain in excluded_domains):
            return False
        
        # 실제 회사 도메인으로 보이는 것들
        company_domains = ['.co.kr', '.com', '.kr', '.org', '.net']
        return any(domain in url_lower for domain in company_domains)
    
    def is_valid_address(self, text):
        """유효한 주소인지 확인"""
        if len(text) < 10 or len(text) > 200:
            return False
        
        # 한국 주소 패턴 확인
        address_keywords = ['시', '도', '구', '군', '동', '로', '길', '번지']
        return any(keyword in text for keyword in address_keywords)
    
    def is_valid_homepage(self, url):
        """유효한 홈페이지 URL인지 확인"""
        if not url.startswith(('http://', 'https://')):
            return False
        
        # 제외할 도메인들 (구인사이트 제외하지 않음 - 나중에 크롤링할 예정)
        excluded_domains = [
            'naver.com', 'google.com', 'daum.net', 'youtube.com',
            'facebook.com', 'instagram.com', 'twitter.com'
        ]
        
        return not any(domain in url.lower() for domain in excluded_domains)
    
    def is_company_website(self, url, company_name):
        """회사 웹사이트인지 확인"""
        if not self.is_valid_homepage(url):
            return False
        
        # 회사명의 영문/한글 부분 추출
        company_parts = re.findall(r'[a-zA-Z]+', company_name.lower())
        
        url_lower = url.lower()
        return any(part in url_lower for part in company_parts if len(part) > 2)
    
    def select_best_homepage(self, homepages, company_name):
        """가장 적합한 홈페이지 선택"""
        if not homepages:
            return ""
        
        # 회사명과 가장 관련성이 높은 것 선택
        scored_homepages = []
        
        for homepage in homepages:
            score = 0
            # .co.kr, .com 도메인 우선
            if '.co.kr' in homepage:
                score += 3
            elif '.com' in homepage:
                score += 2
            
            # 회사명 포함 여부
            company_parts = re.findall(r'[a-zA-Z]+', company_name.lower())
            for part in company_parts:
                if part in homepage.lower():
                    score += 1
            
            scored_homepages.append((score, homepage))
        
        # 점수가 가장 높은 것 선택
        scored_homepages.sort(reverse=True)
        return scored_homepages[0][1]
    
    def update_excel(self):
        """엑셀 파일의 주소, 홈페이지 업데이트"""
        if self.df is None:
            print("엑셀 파일이 로드되지 않았습니다.")
            return False
        
        # 주소, 홈페이지 열이 없으면 생성
        if '주소' not in self.df.columns:
            self.df['주소'] = ""
        if '홈페이지' not in self.df.columns:
            self.df['홈페이지'] = ""
        
        total_companies = len(self.df)
        
        for index, row in self.df.iterrows():
            try:
                # 회사명 가져오기
                company_name = str(row.get('회원사명', '')).strip()
                
                if not company_name or company_name == 'nan':
                    print(f"행 {index + 1}: 회사명이 없습니다.")
                    continue
                
                print(f"진행률: {index + 1}/{total_companies} - {company_name} 검색 중...")
                
                # 이미 정보가 있으면 건너뛰기
                if pd.notna(row.get('주소')) and pd.notna(row.get('홈페이지')):
                    print(f"  -> 이미 정보가 있습니다. 건너뜀")
                    continue
                
                address, homepage = self.search_company_info(company_name)
                
                # 결과 저장
                self.df.at[index, '주소'] = address
                self.df.at[index, '홈페이지'] = homepage
                
                print(f"  -> 주소: {address[:50]}{'...' if len(address) > 50 else ''}")
                print(f"  -> 홈페이지: {homepage}")
                
                # 요청 간 딜레이 (서버 부하 방지)
                time.sleep(2)
                
            except Exception as e:
                print(f"행 {index + 1} 처리 중 오류: {e}")
                continue
        
        return True
    
    def save_excel(self, output_path=None):
        """결과를 엑셀 파일로 저장"""
        if output_path is None:
            output_path = self.excel_file_path.replace('.xlsx', '_업데이트.xlsx')
        
        try:
            self.df.to_excel(output_path, index=False)
            print(f"결과 저장 완료: {output_path}")
            return True
        except Exception as e:
            print(f"파일 저장 실패: {e}")
            return False
    
    def run(self):
        """전체 프로세스 실행"""
        print("회원사 정보 자동 수집을 시작합니다...")
        
        if not self.load_excel():
            return False
        
        if not self.update_excel():
            print("정보 수집 중 오류가 발생했습니다.")
            return False
        
        if not self.save_excel():
            return False
        
        print("모든 작업이 완료되었습니다!")
        return True

# 사용 방법
if __name__ == "__main__":
    # 파일 경로 설정
    excel_file_path = "회원사 목록.xlsx"
    
    # 컬렉터 실행
    collector = CompanyInfoCollector(excel_file_path)
    collector.run()