"""
business_classifier.py
업태명 및 업소명(상호명) 기반 NMIS 업종 자동 분류 모듈
(소분류 / 세분류 / 세세분류)
"""
from __future__ import annotations
import re

def normalize_text(value: str | None) -> str:
    """텍스트 정규화 (공백 제거, 구분자 통일, 소문자화, 전각 괄호 변환)"""
    if value is None:
        return ""
    s = str(value).strip()
    s = re.sub(r'\s+', '', s)
    s = s.replace('·', '/').replace('ㆍ', '/')
    s = s.replace('（', '(').replace('）', ')')
    return s.lower()

# 1. 업태명 정확 일치 / 정규화 일치 기본 규칙
BUSINESS_TYPE_RULES = {
    "한식": {
        "smallCategory": "음식점업",
        "detailCategory": "한식음식점업",
        "subDetailCategory": "한식일반음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "냉면집": {
        "smallCategory": "음식점업",
        "detailCategory": "한식음식점업",
        "subDetailCategory": "한식면요리전문점",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "식육(숯불구이)": {
        "smallCategory": "음식점업",
        "detailCategory": "한식음식점업",
        "subDetailCategory": "한식육류요리전문점",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "횟집": {
        "smallCategory": "음식점업",
        "detailCategory": "한식음식점업",
        "subDetailCategory": "한식해산물요리전문점",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "중국식": {
        "smallCategory": "음식점업",
        "detailCategory": "그외음식점업",
        "subDetailCategory": "중식음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "일식": {
        "smallCategory": "음식점업",
        "detailCategory": "그외음식점업",
        "subDetailCategory": "일식음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "경양식": {
        "smallCategory": "음식점업",
        "detailCategory": "그외음식점업",
        "subDetailCategory": "서양식음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "외국음식전문점(인도,태국등)": {
        "smallCategory": "음식점업",
        "detailCategory": "그외음식점업",
        "subDetailCategory": "기타외국식음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "뷔페식": {
        "smallCategory": "음식점업",
        "detailCategory": "그외음식점업",
        "subDetailCategory": "뷔페음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "분식": {
        "smallCategory": "음식점업",
        "detailCategory": "기타 간이 음식점업",
        "subDetailCategory": "김밥 및 기타 간이 음식점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "호프/통닭": {
        "smallCategory": "음식점업",
        "detailCategory": "기타 간이 음식점업",
        "subDetailCategory": "치킨전문점",
        "confidence": "MEDIUM",
        "reviewRequired": True,
        "reason": "호프/통닭은 치킨전문점과 생맥주 전문점이 혼재할 수 있음"
    },
    "정종/대포집/소주방": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "주점업",
        "subDetailCategory": "기타 주점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "카페": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "비알코올 음료점업",
        "subDetailCategory": "커피전문점",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "까페": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "비알코올 음료점업",
        "subDetailCategory": "커피전문점",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "커피숍": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "비알코올 음료점업",
        "subDetailCategory": "커피전문점",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "다방": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "비알코올 음료점업",
        "subDetailCategory": "기타비알코올 음료점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "유흥주점": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "주점업",
        "subDetailCategory": "일반 유흥주점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "단란주점": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "주점업",
        "subDetailCategory": "일반 유흥주점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "무도유흥주점": {
        "smallCategory": "주점 및 비알코올 음료점업",
        "detailCategory": "주점업",
        "subDetailCategory": "무도 유흥주점업",
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": ""
    },
    "기타": {
        "smallCategory": None,
        "detailCategory": None,
        "subDetailCategory": None,
        "confidence": "LOW",
        "reviewRequired": True,
        "reason": "업태명 기타만으로 업종 판별 불가"
    }
}

# 2. 보조 키워드 매핑 테이블
KEYWORD_CLASSIFICATIONS = [
    # 한식면요리전문점
    {
        "keywords": ["냉면", "막국수", "칼국수", "잔치국수", "비빔국수", "국수", "수제비", "메밀", "면옥"],
        "smallCategory": "음식점업", "detailCategory": "한식음식점업", "subDetailCategory": "한식면요리전문점"
    },
    # 한식육류요리전문점
    {
        "keywords": ["한우", "갈비", "삼겹살", "고깃집", "고기집", "숯불", "식육", "정육식당", "막창", "곱창", "대창", "오리구이", "불고기", "육회"],
        "smallCategory": "음식점업", "detailCategory": "한식음식점업", "subDetailCategory": "한식육류요리전문점"
    },
    # 한식해산물요리전문점
    {
        "keywords": ["횟집", "회센터", "수산", "해산물", "생선회", "아구찜", "해물찜", "조개구이", "장어", "매운탕", "낙지", "문어"],
        "smallCategory": "음식점업", "detailCategory": "한식음식점업", "subDetailCategory": "한식해산물요리전문점"
    },
    # 중식음식점업
    {
        "keywords": ["중국집", "중화요리", "짜장", "짬뽕", "마라탕", "양꼬치"],
        "smallCategory": "음식점업", "detailCategory": "그외음식점업", "subDetailCategory": "중식음식점업"
    },
    # 일식음식점업
    {
        "keywords": ["일식", "초밥", "스시", "돈카츠", "돈까스", "우동", "라멘", "사시미"],
        "smallCategory": "음식점업", "detailCategory": "그외음식점업", "subDetailCategory": "일식음식점업"
    },
    # 서양식음식점업
    {
        "keywords": ["경양식", "레스토랑", "파스타", "스테이크", "브런치"],
        "smallCategory": "음식점업", "detailCategory": "그외음식점업", "subDetailCategory": "서양식음식점업"
    },
    # 기타외국식음식점업
    {
        "keywords": ["인도요리", "태국요리", "베트남요리", "쌀국수", "멕시칸", "터키요리"],
        "smallCategory": "음식점업", "detailCategory": "그외음식점업", "subDetailCategory": "기타외국식음식점업"
    },
    # 뷔페음식점업
    {
        "keywords": ["뷔페", "부페", "buffet"],
        "smallCategory": "음식점업", "detailCategory": "그외음식점업", "subDetailCategory": "뷔페음식점업"
    },
    # 제과점업
    {
        "keywords": ["제과점", "베이커리", "빵집", "제과제빵", "케이크"],
        "smallCategory": "음식점업", "detailCategory": "기타 간이 음식점업", "subDetailCategory": "제과점업"
    },
    # 피자, 햄버거, 샌드위치 및 유사음식점업
    {
        "keywords": ["피자", "햄버거", "버거", "샌드위치", "토스트", "핫도그"],
        "smallCategory": "음식점업", "detailCategory": "기타 간이 음식점업", "subDetailCategory": "피자, 햄버거, 샌드위치 및 유사음식점업"
    },
    # 치킨전문점
    {
        "keywords": ["치킨", "통닭", "닭강정", "후라이드", "양념치킨"],
        "smallCategory": "음식점업", "detailCategory": "기타 간이 음식점업", "subDetailCategory": "치킨전문점"
    },
    # 김밥 및 기타 간이 음식점업
    {
        "keywords": ["분식", "김밥", "떡볶이", "순대", "라볶이", "어묵"],
        "smallCategory": "음식점업", "detailCategory": "기타 간이 음식점업", "subDetailCategory": "김밥 및 기타 간이 음식점업"
    },
    # 간이 음식 포장 판매 전문점
    {
        "keywords": ["포장전문", "테이크아웃전문", "도시락전문", "컵밥", "포장판매"],
        "smallCategory": "음식점업", "detailCategory": "기타 간이 음식점업", "subDetailCategory": "간이 음식 포장 판매 전문점"
    },
    # 커피전문점
    {
        "keywords": ["카페", "까페", "커피", "커피숍", "로스터리", "에스프레소", "coffee", "cafe"],
        "smallCategory": "주점 및 비알코올 음료점업", "detailCategory": "비알코올 음료점업", "subDetailCategory": "커피전문점"
    },
    # 기타비알코올 음료점업
    {
        "keywords": ["다방", "찻집", "전통차", "주스", "생과일주스", "버블티", "음료전문점", "차전문점"],
        "smallCategory": "주점 및 비알코올 음료점업", "detailCategory": "비알코올 음료점업", "subDetailCategory": "기타비알코올 음료점업"
    },
    # 생맥주 전문점
    {
        "keywords": ["호프", "생맥주", "맥주전문점", "비어", "펍", "pub", "beer"],
        "smallCategory": "주점 및 비알코올 음료점업", "detailCategory": "주점업", "subDetailCategory": "생맥주 전문점"
    },
    # 기타 주점업
    {
        "keywords": ["정종", "대포집", "소주방", "포차", "실내포장마차", "주막", "선술집", "술집", "요리주점"],
        "smallCategory": "주점 및 비알코올 음료점업", "detailCategory": "주점업", "subDetailCategory": "기타 주점업"
    },
    # 일반 유흥주점업
    {
        "keywords": ["유흥주점", "단란주점", "룸살롱", "노래주점"],
        "smallCategory": "주점 및 비알코올 음료점업", "detailCategory": "주점업", "subDetailCategory": "일반 유흥주점업"
    },
    # 무도 유흥주점업
    {
        "keywords": ["무도유흥주점", "나이트클럽", "디스코클럽", "무도장"],
        "smallCategory": "주점 및 비알코올 음료점업", "detailCategory": "주점업", "subDetailCategory": "무도 유흥주점업"
    }
]

def classify_business(business_type: str | None, business_name: str = "", item_name: str = "") -> dict:
    """
    업태명, 상호명(업체명), 종목명을 종합 분석하여 소분류/세분류/세세분류를 결정하는 분류 함수
    """
    orig_bt = str(business_type or "").strip()
    norm_bt = normalize_text(orig_bt)
    norm_bn = normalize_text(business_name)
    norm_in = normalize_text(item_name)
    combined_name = f"{norm_bn} {norm_in}".strip()

    result = {
        "originalBusinessType": orig_bt,
        "normalizedBusinessType": norm_bt,
        "smallCategory": None,
        "detailCategory": None,
        "subDetailCategory": None,
        "confidence": "HIGH",
        "reviewRequired": False,
        "reason": "",
        "matchedBy": "EXACT"
    }

    if not norm_bt:
        # 업태명 부재시 키워드 분석 시도
        matched_kw = _match_by_keywords(combined_name)
        if matched_kw:
            result.update(matched_kw)
            result["matchedBy"] = "KEYWORD"
            result["confidence"] = "MEDIUM"
            result["reason"] = f"업태명 없음 -> 상호명 키워드 '{result['matchedKeyword']}' 매칭"
        else:
            result["confidence"] = "LOW"
            result["reviewRequired"] = True
            result["reason"] = "업태명 미입력으로 업종 판별 불가"
            result["matchedBy"] = "UNCLASSIFIED"
        return result

    # 1. 업태명 정확 일치 (EXACT) 또는 정규화/별칭 일치 (ALIAS)
    matched_rule = None
    matched_key = None

    for rule_key, rule_val in BUSINESS_TYPE_RULES.items():
        norm_key = normalize_text(rule_key)
        if orig_bt == rule_key:
            matched_rule = rule_val
            matched_key = rule_key
            result["matchedBy"] = "EXACT"
            break
        elif norm_bt == norm_key:
            matched_rule = rule_val
            matched_key = rule_key
            result["matchedBy"] = "ALIAS"
            break

    # 2. 특별 처리: '호프/통닭' (또는 정규화된 호프/통닭)
    if norm_bt in ("호프/통닭", "호프통닭"):
        chicken_kws = ["치킨", "통닭", "닭강정", "후라이드", "양념치킨", "옛날통닭", "두마리치킨", "닭집", "치킨호프"]
        beer_kws = ["호프", "생맥주", "맥주", "비어", "펍", "pub", "beer"]

        has_chicken = any(kw in combined_name for kw in chicken_kws)
        has_beer = any(kw in combined_name for kw in beer_kws)

        if has_chicken:
            result["smallCategory"] = "음식점업"
            result["detailCategory"] = "기타 간이 음식점업"
            result["subDetailCategory"] = "치킨전문점"
            result["confidence"] = "HIGH"
            result["reviewRequired"] = False
            result["matchedBy"] = "KEYWORD"
            result["reason"] = "호프/통닭 업태 -> 상호명 치킨 관련 키워드 감지"
            return result
        elif has_beer:
            result["smallCategory"] = "주점 및 비알코올 음료점업"
            result["detailCategory"] = "주점업"
            result["subDetailCategory"] = "생맥주 전문점"
            result["confidence"] = "HIGH"
            result["reviewRequired"] = False
            result["matchedBy"] = "KEYWORD"
            result["reason"] = "호프/통닭 업태 -> 상호명 맥주/호프 관련 키워드 감지"
            return result
        else:
            # 치킨/맥주 키워드 모두 없거나 상호명 분해 불가시 기본값(치킨전문점, MEDIUM, reviewRequired: True)
            result["smallCategory"] = "음식점업"
            result["detailCategory"] = "기타 간이 음식점업"
            result["subDetailCategory"] = "치킨전문점"
            result["confidence"] = "MEDIUM"
            result["reviewRequired"] = True
            result["matchedBy"] = "ALIAS"
            result["reason"] = "호프/통닭은 치킨전문점과 생맥주 전문점이 혼재할 수 있음"
            return result

    # 3. 특별 처리: '기타' 업태
    if norm_bt == "기타":
        matched_kw = _match_by_keywords(combined_name)
        if matched_kw:
            result["smallCategory"] = matched_kw["smallCategory"]
            result["detailCategory"] = matched_kw["detailCategory"]
            result["subDetailCategory"] = matched_kw["subDetailCategory"]
            result["confidence"] = "HIGH"
            result["reviewRequired"] = False
            result["matchedBy"] = "KEYWORD"
            result["reason"] = f"기타 업태 -> 상호명 키워드 '{matched_kw['matchedKeyword']}' 매칭"
        else:
            result["smallCategory"] = None
            result["detailCategory"] = None
            result["subDetailCategory"] = None
            result["confidence"] = "LOW"
            result["reviewRequired"] = True
            result["matchedBy"] = "UNCLASSIFIED"
            result["reason"] = "업태명 기타만으로 업종 판별 불가"
        return result

    # 4. 규칙에 일치한 경우 결과 적용
    if matched_rule:
        result["smallCategory"] = matched_rule["smallCategory"]
        result["detailCategory"] = matched_rule["detailCategory"]
        result["subDetailCategory"] = matched_rule["subDetailCategory"]
        result["confidence"] = matched_rule.get("confidence", "HIGH")
        result["reviewRequired"] = matched_rule.get("reviewRequired", False)
        result["reason"] = matched_rule.get("reason", "")
        return result

    # 5. 규칙에 없는 개별 업태 ➔ 상호명 키워드 분석 시도 (KEYWORD)
    matched_kw = _match_by_keywords(norm_bt + " " + combined_name)
    if matched_kw:
        result["smallCategory"] = matched_kw["smallCategory"]
        result["detailCategory"] = matched_kw["detailCategory"]
        result["subDetailCategory"] = matched_kw["subDetailCategory"]
        result["confidence"] = "MEDIUM"
        result["reviewRequired"] = False
        result["matchedBy"] = "KEYWORD"
        result["reason"] = f"키워드 '{matched_kw['matchedKeyword']}' 매칭"
        return result

    # 6. 미분류 (UNCLASSIFIED)
    result["smallCategory"] = None
    result["detailCategory"] = None
    result["subDetailCategory"] = None
    result["confidence"] = "LOW"
    result["reviewRequired"] = True
    result["matchedBy"] = "UNCLASSIFIED"
    result["reason"] = f"업태명 '{orig_bt}'에 대응하는 업종 분류 규칙을 찾을 수 없음"
    return result

def _match_by_keywords(text: str) -> dict | None:
    if not text:
        return None
    for item in KEYWORD_CLASSIFICATIONS:
        for kw in item["keywords"]:
            if kw.lower() in text:
                return {
                    "smallCategory": item["smallCategory"],
                    "detailCategory": item["detailCategory"],
                    "subDetailCategory": item["subDetailCategory"],
                    "matchedKeyword": kw
                }
    return None
