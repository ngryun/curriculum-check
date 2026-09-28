import React, { useState, useMemo, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import { SURVEY_TEMPLATE_COPY_URL } from "./config.js";
import { buildStudentPageHtml, buildAppsScriptCode, buildSurveyConfigCode } from "./survey/appsScript.js";
import { keepRead, keepWrite, keepClear, KEEP_MAX_AGE_DAYS } from "./localKeep.js";

function makeIcon(char) {
  return function Icon({ size = 14, color, style, ...rest }) {
    return (
      <span
        style={{
          fontSize: size,
          lineHeight: 1,
          display: "inline-block",
          color,
          ...style,
        }}
        {...rest}
      >
        {char}
      </span>
    );
  };
}
const ChevronDown = makeIcon("\u25BC");
const ChevronUp = makeIcon("\u25B2");
const Upload = makeIcon("\u2191");
const X = makeIcon("\u2715");
const Download = makeIcon("\u2193");
const Search = makeIcon("\u{1F50D}");
const FileSpreadsheet = makeIcon("\u25A6");
const CheckCircle2 = makeIcon("\u2713");
const AlertCircle = makeIcon("\u26A0\uFE0E");
const Users = makeIcon("\u26A0\uFE0E");
const ArrowRight = makeIcon("\u2192");
const Check = makeIcon("\u2713");
const Info = makeIcon("\u2139\uFE0E");
const Save = makeIcon("\u{1F4BE}");
const FileUp = makeIcon("\u2191");
const Lock = makeIcon("\u{1F512}");

// ---------- Design tokens ----------
const INK = "#1C2333";
const PAPER = "#F6F4EF";
const LINE = "#DDD8CC";
const MUTED = "#6B7280";
const OK = "#2F6D4F";
const OK_BG = "#E7F1EA";
const WARN = "#A2452C";
const WARN_BG = "#F7E9E3";
const ACCENT = "#2C5A8A";
const ACCENT_BG = "#EAF0F6";
const FONT = `"Pretendard Variable", "Apple SD Gothic Neo", "Malgun Gothic", -apple-system, sans-serif`;
// ---------- subject name normalization ----------
// every visually-identical "middle dot" variant seen in real school files:
// · U+00B7, ‧ U+2027, • U+2022, ・ U+30FB, ･ U+FF65, ᐧ U+1427, ⋅ U+22C5 (압핀 '기술⋅가정'), ∙ U+2219, ㆍ U+318D (한글 아래아)
const MIDDLE_DOT_RE = /[·‧•・･ᐧ⋅∙ㆍ]/g;
function normalizeName(s) {
  if (s == null) return "";
  return String(s)
    .trim()
    .replace(/\s+/g, "")
    .replace(/Ⅰ/g, "I")
    .replace(/Ⅱ/g, "II")
    .replace(/Ⅲ/g, "III")
    .replace(/Ⅳ/g, "IV")
    .replace(/Ⅴ/g, "V")
    .replace(MIDDLE_DOT_RE, "·")
    .replace(/[()（）]/g, "");
}
// 교과(군) names are typed by hand into all kinds of Excel templates, and the "middle dot" in names
// like 기술·가정 gets typed with several visually-identical-but-different Unicode characters
// depending on the person's keyboard/IME (see MIDDLE_DOT_RE).
// Silently failing to match these breaks the 생활교양 merge and the national-credit lookup, so every
// group name read from a file is normalized to one canonical dot character right at the source.
function normalizeGroupName(s) {
  if (s == null) return "";
  return String(s).trim().replace(MIDDLE_DOT_RE, "·");
}
// ---------- embedded national data (2022 개정 교육과정) ----------
// subject master list: extracted from the school's own '과목목록_2022개정' reference sheet
// (this is the authoritative, already-structured source — no PDF text-extraction needed)
const NATIONAL_SUBJECT_MASTER = [
  {
    "group": "국어",
    "name": "공통국어1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "공통국어2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "화법과 언어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "독서와 작문",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "문학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "주제 탐구 독서",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "문학과 영상",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "직무 의사소통",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "독서 토론과 글쓰기",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "매체 의사소통",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "국어",
    "name": "언어생활 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "공통수학1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "공통수학2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "기본수학1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "기본수학2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "대수",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "미적분Ⅰ",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "확률과 통계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "기하",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "미적분Ⅱ",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "경제 수학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "인공지능 수학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "직무 수학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "수학과 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "실용 통계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "수학",
    "name": "수학과제 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "공통영어1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "공통영어2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "기본영어1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "기본영어2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "영어Ⅰ",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "영어Ⅱ",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "영어 독해와 작문",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "영미 문학 읽기",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "영어 발표와 토론",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "심화 영어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "심화 영어 독해와 작문",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "직무 영어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "실생활 영어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "미디어 영어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "영어",
    "name": "세계 문화와 영어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "한국사1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "한국사2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "통합사회1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "통합사회2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "세계시민과 지리",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "세계사",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "사회와 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "현대사회와 윤리",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "한국지리 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "도시의 미래 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "동아시아 역사 기행",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "정치",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "법과 사회",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "경제",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "윤리와 사상",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "인문학과 윤리",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "국제 관계의 이해",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "여행지리",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "역사로 탐구하는 현대 세계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "사회문제 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "금융과 경제생활",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "윤리문제 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "사회",
    "name": "기후변화와 지속가능한 세계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "통합과학1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "통합과학2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "과학탐구실험1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "과학탐구실험2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "물리학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "화학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "생명과학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "지구과학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "역학과 에너지",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "전자기와 양자",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "물질과 에너지",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "화학 반응의 세계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "세포와 물질대사",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "생물의 유전",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "지구시스템과학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "행성우주과학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "과학의 역사와 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "기후변화와 환경생태",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학",
    "name": "융합과학 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "체육1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "체육2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "운동과 건강",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "스포츠 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "스포츠 과학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "스포츠 생활1",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "체육",
    "name": "스포츠 생활2",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "음악",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "미술",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "연극",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "음악 연주와 창작",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "음악 감상과 비평",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "미술 창작",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "미술 감상과 비평",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "음악과 미디어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "미술과 매체",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "기술·가정",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "로봇과 공학세계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "생활과학 탐구",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "창의 공학 설계",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "지식 재산 일반",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "생애 설계와 자립",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "아동발달과 부모",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "정보",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "인공지능 기초",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "데이터 과학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "기술·가정/정보",
    "name": "소프트웨어와 생활",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "독일어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "독일어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 독일어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "독일어권 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "프랑스어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "프랑스어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 프랑스어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "프랑스어권 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "스페인어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "스페인어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 스페인어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "스페인어권 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "중국어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "중국어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 중국어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "중국 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "일본어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "일본어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 일본어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "일본 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "러시아어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "러시아어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 러시아어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "러시아 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "아랍어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "아랍어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 아랍어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "아랍 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "베트남어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "베트남어 회화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "심화 베트남어",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "베트남 문화",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "한문",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "한문 고전 읽기",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "제2외국어/한문",
    "name": "언어생활과 한자",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "진로와 직업",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "생태와 환경",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "인간과 철학",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "논리와 사고",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "인간과 심리",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "교육의 이해",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "삶과 종교",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "보건",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "인간과 경제활동",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "논술",
    "type": "일반",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "전문 수학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "이산 수학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 기하",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 대수",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 미적분",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 물리학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 화학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 생명과학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "고급 지구과학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "과학과제 연구",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "물리학 실험",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "화학 실험",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "생명과학 실험",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "지구과학 실험",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "과학계열",
    "name": "정보과학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 개론",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "육상",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "체조",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "수상 스포츠",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "기초 체육 전공 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "심화 체육 전공 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "고급 체육 전공 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 경기 체력",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 경기 기술",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 경기 분석",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 교육",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 생리의학",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "체육계열",
    "name": "스포츠 행정 및 경영",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "음악 이론",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "음악사",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "시창·청음",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "음악 전공 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "합창·합주",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "음악 공연 실습",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "음악과 문화",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "미술 이론",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "드로잉",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "미술사",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "미술 전공 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "조형 탐구",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "미술 매체 탐구",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "미술과 사회",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용의 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용과 몸",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용 기초 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용 전공 실기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "안무",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용 제작 실습",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용 감상과 비평",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무용과 매체",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "문예 창작의 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "문장론",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "문학 감상과 비평",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "시 창작",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "소설 창작",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "극 창작",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "문학과 매체",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "연극과 몸",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "연극과 말",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "연기",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "무대 미술과 기술",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "연극 제작 실습",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "연극 감상과 비평",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "연극과 삶",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "영화의 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "촬영·조명",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "편집·사운드",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "영화 제작 실습",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "영화 감상과 비평",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "영화와 삶",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "사진의 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "사진 촬영",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "사진 표현 기법",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "영상 제작의 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "사진 감상과 비평",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술계열",
    "name": "사진과 삶",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어 독해Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어 독해Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "심화 영어 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "국제 정치",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "국제 경제",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "국제법",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "지역 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "한국 사회의 이해",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "비교 문화",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "세계 문제와 미래 사회",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "국제 관계와 국제기구",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "현대 세계의 변화",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "사회 탐구 방법",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "사회과제 연구",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 독일어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "독일어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "독일어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "독일어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "독일어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 프랑스어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "프랑스어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "프랑스어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "프랑스어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "프랑스어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 스페인어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "스페인어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "스페인어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "스페인어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "스페인어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 중국어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "중국어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "중국어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "중국어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "중국어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 일본어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "일본어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "일본어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "일본어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "일본어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 러시아어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "러시아어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "러시아어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "러시아어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "러시아어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 아랍어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "아랍어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "아랍어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "아랍어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "아랍어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "전공 기초 베트남어",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "베트남어 회화Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "베트남어 회화Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "베트남어 독해와 작문Ⅰ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "외국어·국제계열",
    "name": "베트남어 독해와 작문Ⅱ",
    "type": "전문교과",
    "credit": 4,
  },
  {
    "group": "예술",
    "name": "미술사",
    "type": "진로",
    "credit": 4,
  },
  {
    "group": "교양",
    "name": "R&E",
    "type": "융합",
    "credit": 4,
  },
];
// national minimum required credits per 교과(군), fixed by the 2022 개정 교육과정 총론 —
// used whenever a curriculum file doesn't specify its own (most schools don't store this explicitly)
const NATIONAL_REQUIRED_CREDITS = {
  "국어": 8,
  "수학": 8,
  "영어": 8,
  "사회": 14,
  "과학": 10,
  "체육": 10,
  "예술": 10,
  "생활교양": 16,
};
// the default downloadable 편제표 template file (조사::학교가 실제 사용 중인 입력용 양식), embedded as base64
const TEMPLATE_XLSX_BASE64 =
  "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";
// ---------- sheet helpers ----------
function sheetToMatrix(wb, sheetName) {
  const sheet = wb.Sheets[sheetName];
  if (!sheet) return null;
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: true,
  });
}
function findSheetByFragment(wb, fragment) {
  return wb.SheetNames.find((n) => n.includes(fragment)) || null;
}
async function readWorkbookFromFile(file) {
  const buf = await file.arrayBuffer();
  return XLSX.read(buf, {
    type: "array",
  });
}
// find the sheet that actually holds the row-per-subject curriculum table, regardless of its name/position:
// prefer a sheet literally named with '입력', otherwise pick whichever '구분'-headed sheet has its 교과(군)
// column filled on (almost) every row — the real input table, not a forward-filled display/preview sheet.
function findCourseSheet(wb) {
  const byName = findSheetByFragment(wb, "입력");
  if (byName) return byName;
  let best = null;
  let bestRatio = -1;
  for (const sn of wb.SheetNames) {
    const rows = sheetToMatrix(wb, sn);
    if (!rows) continue;
    const headerIdx = rows.findIndex((r) => r && r[0] === "구분");
    if (headerIdx === -1) continue;
    let total = 0;
    let filled = 0;
    for (let i = headerIdx + 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const gubun = r[0];
      if (!gubun || typeof gubun !== "string") continue;
      if (/^[▣●○■※]/.test(gubun.trim())) continue;
      if (!r[3]) continue;
      total++;
      if (r[1]) filled++;
    }
    if (total === 0) continue;
    const ratio = filled / total;
    if (ratio > bestRatio) {
      bestRatio = ratio;
      best = sn;
    }
  }
  return best || wb.SheetNames[0];
}
// ---------- curriculum parsing ----------
function parseCourseList(wb) {
  const sheetName = findCourseSheet(wb);
  const rows = sheetToMatrix(wb, sheetName);
  if (!rows) throw new Error("편제표입력 시트를 찾을 수 없습니다.");
  const headerIdx = rows.findIndex((r) => r && r[0] === "구분");
  if (headerIdx === -1) throw new Error("‘구분’ 헤더가 있는 표를 찾지 못했습니다. 템플릿 구조를 확인해주세요.");
  const courses = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const gubun = r[0];
    const group = r[1];
    const courseType = r[2];
    const name = r[3];
    const baseCredit = r[4];
    const operCredit = r[5];
    const semester = r[6];
    const selectGroup = r[7];
    if (!gubun || typeof gubun !== "string") continue;
    if (/^[▣●○■※]/.test(gubun.trim())) continue;
    if (!name || !group) continue;
    const oc = toCreditNumber(operCredit);
    const bc = toCreditNumber(baseCredit);
    courses.push({
      gubun: normalizeGubun(gubun),
      group: normalizeGroupName(group),
      courseType: courseType ? String(courseType).trim() : "",
      name: String(name).trim(),
      key: normalizeName(name),
      credit: oc > 0 ? oc : bc > 0 ? bc : 0,
      // 운영학점, 비어 있으면 기준학점
      semester: normalizeSemesterCode(semester),
      selectGroup: selectGroup ? String(selectGroup).trim() : "",
    });
  }
  return courses;
}
function parseGroupRequirements(wb, courses) {
  const req = {};
  // Preferred: an explicit '필수이수학점' column inside 편제표입력 itself (new template) —
  // robust to layout differences across schools since it's read by header name, not position.
  const inputSheetName = findCourseSheet(wb);
  const inputRows = sheetToMatrix(wb, inputSheetName);
  if (inputRows) {
    const headerIdx = inputRows.findIndex((r) => r && r[0] === "구분");
    if (headerIdx !== -1) {
      const header = inputRows[headerIdx];
      const reqCol = header.findIndex((c) => typeof c === "string" && c.includes("필수이수학점"));
      const groupCol = header.findIndex((c) => typeof c === "string" && c.includes("교과"));
      if (reqCol !== -1 && groupCol !== -1) {
        for (let i = headerIdx + 1; i < inputRows.length; i++) {
          const r = inputRows[i];
          if (!r) continue;
          const group = normalizeGroupName(r[groupCol]);
          const val = r[reqCol];
          if (group && val != null && val !== "" && !isNaN(Number(val)) && !(group in req)) {
            req[group] = Number(val);
          }
        }
      }
    }
  }
  // Fallback: derive from the auto-generated '편제표(교과군순)' preview sheet by fixed column position
  // (works with the original template shape that has no explicit column in 편제표입력).
  if (Object.keys(req).length === 0) {
    const sheetName = findSheetByFragment(wb, "교과군순");
    if (sheetName) {
      const rows = sheetToMatrix(wb, sheetName);
      let lastGroup = null;
      let lastReq = null;
      for (const r of rows) {
        if (!r) continue;
        const group = normalizeGroupName(r[1]);
        const name = r[3];
        const reqCredit = r[14];
        if (group) lastGroup = String(group).trim();
        if (reqCredit != null && reqCredit !== "" && !isNaN(Number(reqCredit))) lastReq = Number(reqCredit);
        if (name && lastGroup && lastReq != null && !isNaN(lastReq) && !(lastGroup in req)) {
          req[lastGroup] = lastReq;
        }
      }
      delete req["교과(군)"];
    }
  }
  const groups = Array.from(new Set(courses.map((c) => c.group)));
  groups.forEach((g) => {
    var _a;
    if (!(g in req)) req[g] = (_a = NATIONAL_REQUIRED_CREDITS[g]) !== null && _a !== void 0 ? _a : 0;
  });
  return {
    req,
    groups,
  };
}
function buildNationalSubjectMasterMap() {
  const map = new Map();
  NATIONAL_SUBJECT_MASTER.forEach((s) =>
    map.set(normalizeName(s.name), {
      group: s.group,
      credit: s.credit,
    }),
  );
  return map;
}
function parseSubjectMaster(wb) {
  const sheetName = findSheetByFragment(wb, "과목목록");
  if (!sheetName) return buildNationalSubjectMasterMap();
  const rows = sheetToMatrix(wb, sheetName);
  const headerIdx = rows.findIndex((r) => r && r[0] === "교과(군)");
  if (headerIdx === -1) return buildNationalSubjectMasterMap();
  const map = buildNationalSubjectMasterMap();
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const group = r[0];
    const name = r[1];
    const credit = r[3];
    if (!group || !name) continue;
    if (typeof group === "string" && group.trim().startsWith("★")) continue;
    map.set(normalizeName(name), {
      group: normalizeGroupName(group),
      credit: Number(credit) || 0,
    });
  }
  return map;
}
// 2022 개정 교육과정 상 기술·가정/제2외국어·한문/교양은 실제로는 하나의 '생활·교양' 영역으로 묶여
// 3년간 16학점이 최소기준입니다. 학교 편제표가 이를 별도 교과군으로 나눠뒀더라도 합쳐서 계산합니다.
const LIFESTYLE_MERGE_NAME = "생활교양";
const LIFESTYLE_MERGE_CREDIT = 16;
// 학교마다 '기술·가정/정보', '제2외국어·한문', '기술ㆍ가정', '교양' 등으로 제각각 적기 때문에 구분 기호를 모두 떼고 비교합니다.
// (하나라도 못 알아보면 생활·교양 16학점 기준이 통째로 빠져 학생이 부족해도 '충족'으로 나올 수 있습니다.)
const LIFESTYLE_GROUP_KEYS = new Set([
  "기술가정",
  "기술가정정보",
  "정보",
  "기술가정정보제2외국어한문교양",
  "기술가정제2외국어한문교양",
  "제2외국어한문",
  "제2외국어",
  "한문",
  "교양",
  "생활교양",
]);
function isLifestyleGroup(g) {
  return LIFESTYLE_GROUP_KEYS.has(normalizeGroupName(g).replace(/[\s·/,.\-_()]/g, ""));
}
function toCurriculumGroup(g) {
  const n = normalizeGroupName(g);
  return isLifestyleGroup(n) ? LIFESTYLE_MERGE_NAME : n;
}
function mergeLifestyleGroups(courses, req, groups) {
  const present = groups.filter((g) => isLifestyleGroup(g));
  if (present.length === 0)
    return {
      courses,
      req,
      groups,
    };
  const newCourses = courses.map((c) =>
    isLifestyleGroup(c.group)
      ? {
          ...c,
          group: LIFESTYLE_MERGE_NAME,
        }
      : c,
  );
  const newGroups = [];
  let inserted = false;
  groups.forEach((g) => {
    if (isLifestyleGroup(g)) {
      if (!inserted) {
        newGroups.push(LIFESTYLE_MERGE_NAME);
        inserted = true;
      }
    } else {
      newGroups.push(g);
    }
  });
  const newReq = {
    ...req,
  };
  Object.keys(newReq).forEach((g) => {
    if (isLifestyleGroup(g)) delete newReq[g];
  });
  newReq[LIFESTYLE_MERGE_NAME] = LIFESTYLE_MERGE_CREDIT;
  return {
    courses: newCourses,
    req: newReq,
    groups: newGroups,
  };
}
// ---------- builder page (편제표 직접 작성) helpers ----------
const SEMESTER_CODES = ["1-1", "1-2", "2-1", "2-2", "3-1", "3-2"];
const BUILDER_CATEGORIES = [
  "국어",
  "수학",
  "영어",
  "사회",
  "과학",
  "체육",
  "예술",
  "기술·가정/정보",
  "제2외국어/한문",
  "교양",
  "과학계열",
  "체육계열",
  "예술계열",
  "외국어·국제계열",
];
function emptyBuilderSemester() {
  return {
    mandatory: [],
    pools: [],
  };
}
function emptyBuilderData() {
  const d = {};
  SEMESTER_CODES.forEach((s) => (d[s] = emptyBuilderSemester()));
  return d;
}
// convert the builder's click-based structure into the same flat course-row shape
// parseCourseList() produces, so the rest of the app (page 2/3, exports) needs zero special-casing.
function builderDataToCourses(builderData) {
  const out = [];
  SEMESTER_CODES.forEach((sem) => {
    const s = builderData[sem];
    if (!s) return;
    s.mandatory.forEach((c) => {
      out.push({
        gubun: "학교지정",
        group: c.group,
        courseType: c.type || "",
        name: c.name,
        key: normalizeName(c.name),
        credit: c.credit,
        semester: sem,
        selectGroup: "필수(학교지정)",
      });
    });
    s.pools.forEach((p) => {
      const label = `택${p.requiredN}(${p.requiredN * p.creditPerCourse}학점)`;
      p.courses.forEach((c) => {
        out.push({
          gubun: "학생선택",
          group: c.group,
          courseType: c.type || "",
          name: c.name,
          key: normalizeName(c.name),
          credit: p.creditPerCourse,
          semester: sem,
          selectGroup: label,
        });
      });
    });
  });
  return out;
}
// reverse of builderDataToCourses — lets an uploaded 편제표 file pre-populate the click-based
// '편제표 직접 작성' page, so the two entry points always stay in sync.
function coursesToBuilderData(rawCourses) {
  const data = emptyBuilderData();
  const poolIndex = new Map(); // `${sem}||${selectGroup}` -> pool object
  rawCourses.forEach((c) => {
    var _a;
    const sem = SEMESTER_CODES.includes(c.semester) ? c.semester : null;
    if (!sem) return; // can't place a course with an unrecognized semester code
    if (c.gubun === "학교지정") {
      data[sem].mandatory.push({
        name: c.name,
        group: c.group,
        type: c.courseType,
        credit: c.credit,
      });
      return;
    }
    const poolKey = `${sem}||${c.selectGroup || "택1"}`;
    if (!poolIndex.has(poolKey)) {
      const requiredN = (_a = parseRequiredN(c.selectGroup)) !== null && _a !== void 0 ? _a : 1;
      const pool = {
        id: poolKey,
        requiredN,
        creditPerCourse: c.credit,
        courses: [],
      };
      poolIndex.set(poolKey, pool);
      data[sem].pools.push(pool);
    }
    poolIndex.get(poolKey).courses.push({
      name: c.name,
      group: c.group,
      type: c.courseType,
    });
  });
  return data;
}
// 학생이 신청한 (과목, 학기)를 편제표 과목에 연결합니다.
// 예전에는 학기가 달라도 같은 이름의 첫 과목으로 조용히 대신해서, 3-1에 신청한 과목이 2-2 과목으로 표시되는
// 오류가 있었습니다. 이제 학기를 아는 경우에는 같은 학기 과목만 그대로 인정하고, 이름만 같고 학기가 다르면
// mismatch로 따로 알려줍니다 (학점·교과군은 반영하되 학기는 신청한 학기로 표시, 택N 계산에서는 제외).
function resolveTakenCourse(list, semester) {
  if (!list || list.length === 0)
    return {
      course: null,
      mismatch: false,
    };
  if (!semester)
    return {
      course: list[0],
      mismatch: false,
    };
  const exact = list.find((c) => c.semester === semester);
  if (exact)
    return {
      course: exact,
      mismatch: false,
    };
  return {
    course: list[0],
    mismatch: true,
  };
}
function nationalOnlyRequirements(courses) {
  const groups = Array.from(new Set(courses.map((c) => c.group)));
  const req = {};
  groups.forEach((g) => {
    var _a;
    return (req[g] = (_a = NATIONAL_REQUIRED_CREDITS[g]) !== null && _a !== void 0 ? _a : 0);
  });
  return {
    req,
    groups,
  };
}
function buildBuilderXlsxBytes(builderData) {
  const header = [
    "구분",
    "교과(군)",
    "과목구분",
    "과목명",
    "기준학점",
    "운영학점",
    "개설학기",
    "선택구분(택N)",
    "비고",
  ];
  const rows = [header];
  SEMESTER_CODES.forEach((sem) => {
    const s = builderData[sem];
    if (!s || (s.mandatory.length === 0 && s.pools.length === 0)) return;
    rows.push([`▣ ${formatSemester(sem)}`]);
    if (s.mandatory.length > 0) {
      rows.push(["● 학교지정 (전교생 공통이수)"]);
      s.mandatory.forEach((c) => {
        rows.push(["학교지정", c.group, c.type || "", c.name, c.credit, c.credit, sem, "필수(학교지정)", ""]);
      });
    }
    s.pools.forEach((p) => {
      const label = `택${p.requiredN}(${p.requiredN * p.creditPerCourse}학점)`;
      rows.push([`○ 학생선택 — ${label}`]);
      p.courses.forEach((c) => {
        rows.push(["학생선택", c.group, c.type || "", c.name, p.creditPerCourse, p.creditPerCourse, sem, label, ""]);
      });
    });
  });
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = header.map(() => ({
    wch: 16,
  }));
  XLSX.utils.book_append_sheet(wb, ws, "편제표입력");
  return xlsxBytesWithTextCells(wb);
}
function exportBuilderToExcel(builderData) {
  downloadBlob(
    new Blob([buildBuilderXlsxBytes(builderData)], {
      type: "application/octet-stream",
    }),
    "편제표_직접작성.xlsx",
  );
}
// Shape A: multi-row header with explicit '이름' / '과목' / '고유번호' labels, grade cols hold the student's numeric ID.
// Shape B: single header row with '학년/반/번호/성명' style labels; a credit row sits directly below/above the name row.
// A single sheet can contain MULTIPLE stacked tables (e.g. a 3-1 block followed by a separate 3-2 block) —
// each header occurrence starts its own block, and each block's data rows stop right before the next block starts.
function decodeIdFromDigits(idKey) {
  const m = /^(\d)(\d{2})(\d{2})$/.exec(String(idKey || "").trim());
  if (!m) return null;
  return {
    grade: Number(m[1]),
    classNum: Number(m[2]),
    number: Number(m[3]),
  };
}
// ---------- class-reassignment mapping (반 편성 발표 자료), used to auto-link 동명이인 across grades ----------
// Accepts basically any layout that has a 성명/이름 column plus a pair of (반, 번호) columns tagged
// with two different grade numbers (e.g. "1학년 반" / "2학년 반") — column order and header wording
// don't matter, only which grade number and 반/번호 keyword each header carries.
function parseClassMappingFile(wb) {
  var _a, _b, _c, _d;
  const sheetName = wb.SheetNames[0];
  const rows = sheetToMatrix(wb, sheetName);
  if (!rows) throw new Error("시트를 읽을 수 없습니다.");
  let headerRow = -1;
  let nameCol = -1;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const r = rows[i];
    if (!r) continue;
    const idx = r.findIndex(
      (c) => typeof c === "string" && ["성명", "이름", "학생성명", "학생이름"].includes(c.trim()),
    );
    if (idx !== -1) {
      headerRow = i;
      nameCol = idx;
      break;
    }
  }
  if (headerRow === -1) throw new Error("‘성명’ 또는 ‘이름’ 열을 찾지 못했습니다.");
  const header = rows[headerRow];
  const cols = header
    .map((h, i) => {
      if (i === nameCol || h == null) return null;
      const s = String(h);
      const gradeMatch = /([123])\s*학년/.exec(s);
      const isNum = /번호|번/.test(s);
      const isClass = /반/.test(s) && !isNum;
      if (!isClass && !isNum) return null;
      return {
        i,
        grade: gradeMatch ? Number(gradeMatch[1]) : null,
        isClass,
        isNum,
      };
    })
    .filter(Boolean);
  const grades = Array.from(new Set(cols.map((c) => c.grade).filter((g) => g != null))).sort((a, b) => a - b);
  if (grades.length >= 2) {
    // ---- Strategy A: separate '반'/'번호' columns tagged with a grade number ----
    const beforeGrade = grades[0];
    const afterGrade = grades[grades.length - 1];
    const beforeClassCol =
      (_a = cols.find((c) => c.grade === beforeGrade && c.isClass)) === null || _a === void 0 ? void 0 : _a.i;
    const beforeNumCol =
      (_b = cols.find((c) => c.grade === beforeGrade && c.isNum)) === null || _b === void 0 ? void 0 : _b.i;
    const afterClassCol =
      (_c = cols.find((c) => c.grade === afterGrade && c.isClass)) === null || _c === void 0 ? void 0 : _c.i;
    const afterNumCol =
      (_d = cols.find((c) => c.grade === afterGrade && c.isNum)) === null || _d === void 0 ? void 0 : _d.i;
    if (beforeClassCol != null && beforeNumCol != null && afterClassCol != null && afterNumCol != null) {
      const entries = [];
      for (let i = headerRow + 1; i < rows.length; i++) {
        const r = rows[i];
        if (!r) continue;
        const name = r[nameCol];
        if (!name || typeof name !== "string" || !name.trim()) continue;
        const bC = r[beforeClassCol];
        const bN = r[beforeNumCol];
        const aC = r[afterClassCol];
        const aN = r[afterNumCol];
        if ((aC == null || aC === "") && (aN == null || aN === "")) continue; // per-file convention: blank after-side = skip
        if (bC == null || bC === "" || bN == null || bN === "") continue;
        entries.push({
          name: name.trim(),
          from: {
            grade: beforeGrade,
            classNum: Number(bC),
            number: Number(bN),
          },
          to: {
            grade: afterGrade,
            classNum: Number(aC),
            number: Number(aN),
          },
        });
      }
      if (entries.length > 0) return entries;
    }
  }
  // ---- Strategy B: fallback — two columns of combined 5-digit '학번' codes (학년+반+번호, e.g. "10305"),
  // one for before and one for after, distinguished by their decoded grade digit ----
  const dataRows = rows.slice(headerRow + 1).filter((r) => r && r[nameCol]);
  const candidateCols = [];
  for (let c = 0; c < header.length; c++) {
    if (c === nameCol) continue;
    let hits = 0;
    let total = 0;
    dataRows.forEach((r) => {
      const v = r[c];
      if (v == null || v === "") return;
      total++;
      if (decodeIdFromDigits(v)) hits++;
    });
    if (total > 0 && hits / total >= 0.8) candidateCols.push(c);
  }
  if (candidateCols.length >= 2) {
    // pick the two candidate columns whose decoded grades differ the most consistently
    const [colA, colB] = candidateCols;
    const entries = [];
    for (const r of dataRows) {
      const name = r[nameCol];
      const a = decodeIdFromDigits(r[colA]);
      const b = decodeIdFromDigits(r[colB]);
      if (!a || !b) continue;
      const [from, to] = a.grade <= b.grade ? [a, b] : [b, a];
      entries.push({
        name: String(name).trim(),
        from,
        to,
      });
    }
    if (entries.length > 0) return entries;
    throw new Error("학번 코드(예: 10305)를 찾았지만 학년 구분이 서로 같습니다. 변경전·변경후 학번 열을 확인해주세요.");
  }
  throw new Error(
    "형식을 인식하지 못했습니다. '1학년 반', '1학년 번호', '2학년 반', '2학년 번호' 같은 열 이름을 쓰거나, 10305(1학년 3반 5번) 같은 5자리 학번 코드를 두 열(변경 전/후)에 넣어주세요.",
  );
}
function idKeyOf(name, f) {
  if (!f || f.grade == null || f.classNum == null || f.number == null) return null;
  return `${name}||${f.grade}-${f.classNum}-${f.number}`;
}
// simple union-find over the (name, grade, class, number) identity keys touched by any mapping entry
function buildClassMappingIndex(entries) {
  const parent = new Map();
  function find(k) {
    let root = k;
    while (parent.has(root) && parent.get(root) !== root) root = parent.get(root);
    let cur = k;
    while (parent.has(cur) && parent.get(cur) !== root) {
      const next = parent.get(cur);
      parent.set(cur, root);
      cur = next;
    }
    return root;
  }
  function union(a, b) {
    if (!parent.has(a)) parent.set(a, a);
    if (!parent.has(b)) parent.set(b, b);
    const ra = find(a);
    const rb = find(b);
    if (ra !== rb) parent.set(ra, rb);
  }
  entries.forEach((e) => {
    const k1 = idKeyOf(e.name, e.from);
    const k2 = idKeyOf(e.name, e.to);
    if (k1 && k2) union(k1, k2);
  });
  return {
    has: (k) => parent.has(k),
    find: (k) => find(k),
  };
}
// ---- 고교학점제 수강신청 누리집 양식 (Shape A) — returns null when the sheet isn't this format ----
function parseHakjeomRows(rows) {
  // find ALL rows containing an exact '이름' cell (Shape A blocks)
  const nameRowsExact = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const idx = r.findIndex((c) => c === "이름");
    if (idx !== -1)
      nameRowsExact.push({
        row: i,
        col: idx,
      });
  }
  if (nameRowsExact.length > 0) {
    // ---- Shape A (one or more stacked blocks) ----
    const records = [];
    nameRowsExact.forEach((block, blockIdx) => {
      var _a;
      const nameRowExact = block.row;
      const nameCol = block.col;
      const blockEnd = blockIdx + 1 < nameRowsExact.length ? nameRowsExact[blockIdx + 1].row : rows.length;
      const labelCol = nameCol + 1;
      const dataStartCol = labelCol + 1;
      let subjectRow = -1;
      let idRow = -1;
      for (let i = nameRowExact + 1; i < Math.min(blockEnd, nameRowExact + 10); i++) {
        const r = rows[i];
        if (!r) continue;
        if (r[labelCol] === "과목") subjectRow = i;
        if (r[labelCol] === "고유번호") idRow = i;
      }
      if (subjectRow === -1) return; // not a real block, skip
      const dataStartRow = (idRow !== -1 ? idRow : subjectRow) + 1;
      const subjectRowData = rows[subjectRow];
      // the header row itself (nameRowExact) sparsely marks semester codes like "3-1"/"3-2" at the
      // start of each column range — forward-fill them so every subject column knows its own semester.
      const headerRowData = rows[nameRowExact];
      // 이름 왼쪽의 학번 열 머리글(‘1학년’/‘2학년’/‘3학년’) — 어느 학년 학번인지 알아 둡니다
      const gradeByCol = {};
      for (let c = 0; c < nameCol; c++) {
        const m = /([123])\s*학년/.exec(String((_a = headerRowData[c]) !== null && _a !== void 0 ? _a : ""));
        if (m) gradeByCol[c] = Number(m[1]);
      }
      const semesterByCol = {};
      let currentSem = null;
      for (let c = dataStartCol; c < Math.max(subjectRowData.length, headerRowData.length); c++) {
        const v = headerRowData[c];
        if (v != null && /^\d-\d$/.test(String(v).trim())) currentSem = String(v).trim();
        semesterByCol[c] = currentSem;
      }
      const subjectCols = [];
      for (let c = dataStartCol; c < subjectRowData.length; c++) {
        if (subjectRowData[c] != null && subjectRowData[c] !== "") {
          subjectCols.push({
            col: c,
            name: String(subjectRowData[c]).trim(),
            semester: semesterByCol[c] || null,
          });
        }
      }
      for (let i = dataStartRow; i < blockEnd; i++) {
        const r = rows[i];
        if (!r) continue;
        const name = r[nameCol];
        if (!name || typeof name !== "string" || !name.trim()) continue;
        const ids = hakjeomIdsOfRow(r, nameCol, gradeByCol);
        if (ids.length === 0) continue;
        const idVal = ids[ids.length - 1].value; // 가장 높은 학년의 학번 = 지금 학번
        const pastIds = ids
          .slice(0, -1)
          .filter((x) => x.fields)
          .map((x) => x.fields);
        const subjects = [];
        for (const sc of subjectCols) {
          if (r[sc.col] === 1 || r[sc.col] === "1")
            subjects.push({
              name: sc.name,
              semester: sc.semester,
            });
        }
        const semesters = Array.from(new Set(subjects.map((s) => s.semester).filter(Boolean)));
        records.push({
          idText: pastIds.length
            ? `학번 ${idVal} (${pastIds.map((p) => `${p.grade}학년 때 ${p.grade}${String(p.classNum).padStart(2, "0")}${String(p.number).padStart(2, "0")}`).join(", ")})`
            : `학번 ${idVal}`,
          idKey: idVal,
          idFields: decodeIdFromDigits(idVal),
          pastIds,
          name: name.trim(),
          subjects,
          semesters,
          blockId: `A${blockIdx}`,
        });
      }
    });
    return records.length > 0 ? records : null;
  }
  return null;
}
// 누리집 파일의 학번 칸: 머리글이 ‘1학년/2학년/3학년’이면 학년별로 읽어 학년 순으로 돌려줍니다(마지막 = 지금 학번).
// 보통은 지금 학년 칸만 채워져 있지만(예: 1학년 0 · 2학년 20101 · 3학년 0), 1학년 칸에 10101처럼 1학년 때 학번이
// 함께 있으면 그 학번도 남겨 동명이인을 자동으로 연결합니다. 머리글이 없는 파일은 예전처럼 첫 번째 학번 하나만 씁니다.
function hakjeomIdsOfRow(r, nameCol, gradeByCol) {
  const out = [];
  for (let c = 0; c < nameCol; c++) {
    const v = r[c];
    if (v == null || v === "" || isNaN(Number(v)) || Number(v) === 0) continue;
    const value = String(v).trim();
    const fields = decodeIdFromDigits(value);
    out.push({
      value,
      fields,
      grade: gradeByCol[c] || (fields && fields.grade) || null,
    });
  }
  if (Object.keys(gradeByCol || {}).length === 0) return out.slice(0, 1);
  return out.filter((x) => x.grade).sort((a, b) => a.grade - b.grade);
}
// 누리집 파일에 적힌 ‘1학년 학번 → 지금 학번’ 쌍은 학급 편성 변경 자료와 같은 역할을 합니다.
function hakjeomMappingEntries(regFiles) {
  const out = [];
  (regFiles || []).forEach((f) => {
    if (f.status !== "ok" || f.disabled) return;
    (f.records || []).forEach((r) => {
      if (!r || !r.idFields || !Array.isArray(r.pastIds)) return;
      r.pastIds.forEach((p) => {
        if (p && p.grade && p.classNum && p.number && p.grade !== r.idFields.grade)
          out.push({
            name: r.name,
            from: p,
            to: r.idFields,
          });
      });
    });
  });
  return out;
}
// 학생 묶기 규칙(합치기 > 편성 자료·누리집 학번 연결 > 이름+학번 > 이름). 화면과 검사 코드가 같은 규칙을 씁니다.
function makeStudentKeyFns(classMappingIndex, ambiguousNames, uidToMergeGroup) {
  function identityKeyFor(rec) {
    const k = idKeyOf(rec.name, rec.idFields);
    if (k && classMappingIndex.has(k)) return "mapped:" + classMappingIndex.find(k);
    if (k) return "id:" + k;
    return "solo:" + rec.uid; // no usable 학번 -> treated as its own person
  }
  function groupKeyFor(rec) {
    if (uidToMergeGroup.has(rec.uid)) return "merge:" + uidToMergeGroup.get(rec.uid);
    if (ambiguousNames.has(rec.name)) return identityKeyFor(rec);
    return "name:" + rec.name;
  }
  return {
    identityKeyFor,
    groupKeyFor,
  };
}
// ---- 압핀(수강신청 프로그램) 결과 양식 (Shape B) — returns null when the sheet isn't this format ----
function parseApinRows(rows) {
  // look for ALL rows with a '성명' fragment cell (one or more stacked blocks)
  const metaRows = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r) continue;
    const idx = r.findIndex((c) => typeof c === "string" && c.includes("성명"));
    if (idx !== -1)
      metaRows.push({
        row: i,
        col: idx,
      });
  }
  if (metaRows.length === 0) return null;
  const records = [];
  metaRows.forEach((block, blockIdx) => {
    const metaRow = block.row;
    const nameCol2 = block.col;
    // stop one row early: the next block's own subject-name row sits directly above its metaRow
    const blockEnd = blockIdx + 1 < metaRows.length ? metaRows[blockIdx + 1].row - 1 : rows.length;
    const idColDefs = [];
    for (let c = 0; c < nameCol2; c++) {
      const label = rows[metaRow][c];
      idColDefs.push({
        col: c,
        label: label != null ? String(label).trim() : `열${c + 1}`,
      });
    }
    const subjectRow2 = metaRow - 1;
    const subjectRowData2 = rows[subjectRow2] || [];
    const subjectCols2 = [];
    for (let c = nameCol2 + 1; c < Math.max(subjectRowData2.length, rows[metaRow].length); c++) {
      const sName = subjectRowData2[c];
      if (sName != null && String(sName).trim() !== "") {
        subjectCols2.push({
          col: c,
          name: String(sName).trim(),
        });
      }
    }
    if (subjectCols2.length === 0) return; // not a real block, skip
    const dataStartRow2 = metaRow + 1;
    for (let i = dataStartRow2; i < blockEnd; i++) {
      const r = rows[i];
      if (!r) continue;
      const name = r[nameCol2];
      if (!name || typeof name !== "string" || !name.trim()) continue;
      const idParts = idColDefs
        .map((d) => {
          var _a;
          return `${(_a = r[d.col]) !== null && _a !== void 0 ? _a : ""}${d.label}`;
        })
        .join(" ");
      const idKey = idColDefs
        .map((d) => {
          var _a;
          return (_a = r[d.col]) !== null && _a !== void 0 ? _a : "";
        })
        .join("-");
      let idFields = null;
      const gradeCol = idColDefs.find((d) => d.label.includes("학년"));
      const classCol = idColDefs.find((d) => d.label.includes("반"));
      const numCol = idColDefs.find((d) => d.label.includes("번호"));
      if (gradeCol && classCol && numCol) {
        idFields = {
          grade: Number(r[gradeCol.col]) || null,
          classNum: Number(r[classCol.col]) || null,
          number: Number(r[numCol.col]) || null,
        };
      }
      const subjects = [];
      for (const sc of subjectCols2) {
        if (r[sc.col] === 1 || r[sc.col] === "1")
          subjects.push({
            name: sc.name,
            semester: null,
          });
      }
      const idText =
        idFields && idFields.grade && idFields.classNum && idFields.number
          ? `${idFields.grade}학년 ${idFields.classNum}반 ${idFields.number}번`
          : idParts.trim() || "-";
      records.push({
        idText,
        idKey,
        idFields,
        name: name.trim(),
        subjects,
        semesters: [],
        blockId: `B${blockIdx}`,
      });
    }
  });
  return records.length > 0 ? records : null;
}
// ---------- 수강신청 파일의 두 가지 출처 ----------
// 두 양식을 한 칸에 섞어 올리면 학기·학번 해석이 서로 꼬여서(압핀은 학기 정보가 파일 안에 없음) 학생에게
// 신청하지 않은 과목이 붙거나 엉뚱한 학기로 표시되는 문제가 있었습니다. 그래서 출처별로 칸을 나누고,
// 파일 내용으로 양식을 다시 확인해서 다른 칸의 파일이면 계산에 넣지 않고 옮기기를 안내합니다.
const REG_SOURCES = {
  hakjeom: {
    id: "hakjeom",
    badge: "A",
    short: "고교학점제",
    color: ACCENT,
    title: "고교학점제 수강신청 누리집 파일",
    desc: "고교학점제 누리집(수강신청 시스템)에서 내려받은 결과 엑셀입니다. 한 파일에 3-1·3-2처럼 여러 학기가 함께 들어 있어도 되고, 학기는 파일 안의 머리글에서 자동으로 읽습니다.",
    signature: "‘이름 · 학기 · 선택그룹 · 과목 · 고유번호’ 머리글 행이 있고, 신청 여부가 1/0으로 적혀 있음",
    dropLabel: "고교학점제 누리집 파일을 끌어다 놓거나",
  },
  apin: {
    id: "apin",
    badge: "B",
    short: "압핀",
    color: "#6B4C9A",
    title: "압핀 수강신청 결과 파일",
    desc: "압핀(수강신청 프로그램)에서 내려받은 결과 엑셀입니다. 압핀 파일은 한 파일이 한 학기라서 학기별로 따로 올려주세요. 학기는 파일 이름(예: ‘2학년 1학기’)에서 읽고, 없으면 편제표와 과목 구성을 비교해 추정합니다.",
    signature: "‘학년 · 반 · 번호 · 성명/단위’ 머리글이고, 과목명 바로 아래 줄에 학점(단위)이 있음",
    dropLabel: "압핀 결과 파일을 끌어다 놓거나",
  },
};
function parseRegistrationAs(wb, source) {
  const preferred = source === "hakjeom" ? "수강신청" : "과목선택";
  const sheetNames = [...wb.SheetNames].sort(
    (a, b) => (b.includes(preferred) ? 1 : 0) - (a.includes(preferred) ? 1 : 0),
  );
  for (const sn of sheetNames) {
    const rows = sheetToMatrix(wb, sn);
    if (!rows) continue;
    const records = source === "hakjeom" ? parseHakjeomRows(rows) : parseApinRows(rows);
    if (records && records.length > 0) return records;
  }
  return null;
}
// 고교학점제 양식을 먼저 검사합니다 — '이름' 아래 '과목'/'고유번호' 머리글이 훨씬 구체적이라 오탐이 거의 없고,
// 압핀 양식('성명' 머리글)은 그 다음에 봅니다.
function detectRegistrationFormat(wb) {
  const hakjeom = parseRegistrationAs(wb, "hakjeom");
  if (hakjeom)
    return {
      source: "hakjeom",
      records: hakjeom,
    };
  const apin = parseRegistrationAs(wb, "apin");
  if (apin)
    return {
      source: "apin",
      records: apin,
    };
  return null;
}
// 파일에 들어 있는 과목명 묶음이 편제표의 어느 학기와 가장 잘 맞는지로 학기를 추정합니다
// (파일 이름에 학기가 없거나, 파일 이름과 실제 내용이 어긋날 때 알려주는 데 사용).
function guessSemesterFromSubjects(subjectNames, courses) {
  const keys = Array.from(new Set(subjectNames.map(normalizeName))).filter(Boolean);
  if (keys.length === 0 || !courses || courses.length === 0) return null;
  const bySem = new Map();
  courses.forEach((c) => {
    if (!SEMESTER_CODES.includes(c.semester)) return;
    if (!bySem.has(c.semester)) bySem.set(c.semester, new Set());
    bySem.get(c.semester).add(c.key);
  });
  const scored = Array.from(bySem.entries())
    .map(([sem, set]) => ({
      sem,
      hits: keys.filter((k) => set.has(k)).length,
    }))
    .sort((a, b) => b.hits - a.hits);
  if (scored.length === 0 || scored[0].hits === 0) return null;
  const best = scored[0];
  if (scored[1] && scored[1].hits === best.hits) return null; // 동점이면 추정하지 않음
  if (best.hits / keys.length < 0.6) return null;
  return {
    sem: best.sem,
    hits: best.hits,
    total: keys.length,
  };
}
// 파일 하나를 편제표와 (과목명 + 학기) 기준으로 대조합니다 — ② 화면에서 파일마다 무엇이 어긋나는지 바로 보여줍니다.
function analyzeRegFile(records, fileSemester, courses) {
  const courseMap = new Map();
  courses.forEach((c) => {
    if (!courseMap.has(c.key)) courseMap.set(c.key, []);
    courseMap.get(c.key).push(c);
  });
  const entries = new Map();
  records.forEach((r) =>
    r.subjects.forEach((s) => {
      const sem = s.semester || fileSemester || null;
      const k = `${normalizeName(s.name)}||${sem || ""}`;
      if (!entries.has(k))
        entries.set(k, {
          name: s.name,
          sem,
          count: 0,
        });
      entries.get(k).count += 1;
    }),
  );
  const matched = [];
  const semMismatch = [];
  const notInCurriculum = [];
  entries.forEach((e) => {
    const list = courseMap.get(normalizeName(e.name));
    if (!list) notInCurriculum.push(e);
    else if (e.sem && !list.some((c) => c.semester === e.sem))
      semMismatch.push({
        ...e,
        offered: Array.from(new Set(list.map((c) => c.semester).filter(Boolean))),
      });
    else matched.push(e);
  });
  return {
    total: entries.size,
    matched,
    semMismatch,
    notInCurriculum,
  };
}
// ---------- 내려받는 엑셀 양식: 모든 칸을 셀 서식 '텍스트'로 ----------
// 엑셀은 '1-1'을 입력하면 날짜(1월 1일)로 바꿔 버립니다. 내려받는 양식은 이미 적힌 칸뿐 아니라 빈칸·새로 적을 줄까지
// 모두 셀 서식 '텍스트'(@)로 해 두어, 선생님이 셀 서식을 따로 바꾸지 않아도 적은 그대로 남게 합니다.
// 이미 숫자로 들어 있는 학점·반·번호도 글자로 바꾸고(프로그램은 글자로 된 숫자도 그대로 읽습니다),
// '텍스트 형식으로 저장된 숫자' 초록 삼각형 경고는 꺼 둡니다. xlsx 안의 XML만 고치므로 글꼴·색·병합·틀 고정은 그대로입니다.
const TEXT_NUMFMT_ID = 49;
const XLSX_MAX_COL = 16384;
function textFormatStylesXml(xml) {
  const m = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(xml);
  if (!m) throw new Error("styles.xml에 cellXfs가 없습니다.");
  const xfs = m[1].match(/<xf\b[^>]*?(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
  if (xfs.length === 0) throw new Error("styles.xml에 셀 서식이 없습니다.");
  // every existing style gets a twin that differs only by the number format (text); twin index = original + offset
  const twins = xfs.map((xf) => {
    const head = /^<xf\b[^>]*?(?=\/?>)/.exec(xf)[0];
    let h = /\snumFmtId="\d+"/.test(head)
      ? head.replace(/\snumFmtId="\d+"/, ` numFmtId="${TEXT_NUMFMT_ID}"`)
      : `${head} numFmtId="${TEXT_NUMFMT_ID}"`;
    h = /\sapplyNumberFormat="[^"]*"/.test(h)
      ? h.replace(/\sapplyNumberFormat="[^"]*"/, ' applyNumberFormat="1"')
      : `${h} applyNumberFormat="1"`;
    return h + xf.slice(head.length);
  });
  const block = `<cellXfs count="${xfs.length * 2}">${xfs.join("")}${twins.join("")}</cellXfs>`;
  return {
    xml: xml.slice(0, m.index) + block + xml.slice(m.index + m[0].length),
    offset: xfs.length,
  };
}
function textFormatSheetXml(xml, offset) {
  const textStyle = (s) => (Number(s) || 0) + offset;
  // cells: point at the text twin of their style; plain numbers become inline text
  let body = xml.replace(/<c\b([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, (all, attrs, tail, inner) => {
    const a = /\ss="\d+"/.test(attrs)
      ? attrs.replace(/\ss="(\d+)"/, (_, s) => ` s="${textStyle(s)}"`)
      : `${attrs} s="${textStyle(0)}"`;
    const t = /\st="([^"]*)"/.exec(attrs);
    if (inner != null && (!t || t[1] === "n") && !/<f[\s>]/.test(inner)) {
      const v = /<v>([^<]*)<\/v>/.exec(inner);
      if (v) return `<c${a.replace(/\st="[^"]*"/, "")} t="inlineStr"><is><t>${v[1]}</t></is></c>`;
    }
    return `<c${a}${tail}`;
  });
  body = body.replace(/<row\b([^>]*)>/g, (all, attrs) =>
    /\scustomFormat="(1|true)"/.test(attrs)
      ? `<row${attrs.replace(/\ss="(\d+)"/, (_, s) => ` s="${textStyle(s)}"`)}>`
      : all,
  );
  // columns: every column from A to the last one defaults to text, so cells typed later in empty rows stay text too
  const defaultWidth = (/<sheetFormatPr\b[^>]*\sdefaultColWidth="([\d.]+)"/.exec(body) || [])[1] || "9.140625";
  const colsMatch = /<cols>([\s\S]*?)<\/cols>/.exec(body);
  const ranges = (colsMatch ? colsMatch[1].match(/<col\b[^>]*\/>/g) || [] : [])
    .map((el) => ({
      el,
      min: Number((/\smin="(\d+)"/.exec(el) || [])[1]),
      max: Number((/\smax="(\d+)"/.exec(el) || [])[1]),
    }))
    .filter((r) => r.min >= 1 && r.max >= r.min)
    .sort((x, y) => x.min - y.min);
  const gap = (from, to) => `<col min="${from}" max="${to}" width="${defaultWidth}" style="${textStyle(0)}"/>`;
  const parts = [];
  let next = 1;
  ranges.forEach((r) => {
    if (r.min < next) return;
    if (r.min > next) parts.push(gap(next, r.min - 1));
    const s = /\sstyle="(\d+)"/.exec(r.el);
    parts.push(
      s
        ? r.el.replace(/\sstyle="\d+"/, ` style="${textStyle(s[1])}"`)
        : r.el.replace(/\s*\/>$/, ` style="${textStyle(0)}"/>`),
    );
    next = Math.min(r.max, XLSX_MAX_COL) + 1;
  });
  if (next <= XLSX_MAX_COL) parts.push(gap(next, XLSX_MAX_COL));
  const colsXml = `<cols>${parts.join("")}</cols>`;
  body = colsMatch
    ? body.slice(0, colsMatch.index) + colsXml + body.slice(colsMatch.index + colsMatch[0].length)
    : body.replace("<sheetData", colsXml + "<sheetData");
  // no green 'number stored as text' triangles (ignoredErrors must sit before drawing/tableParts/extLst)
  body = body.replace(/<ignoredErrors>[\s\S]*?<\/ignoredErrors>|<ignoredErrors\s*\/>/g, "");
  const ignored = '<ignoredErrors><ignoredError sqref="A1:XFD1048576" numberStoredAsText="1"/></ignoredErrors>';
  const later =
    /<(smartTags|drawing|legacyDrawing|legacyDrawingHF|drawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(
      body,
    );
  return later
    ? body.slice(0, later.index) + ignored + body.slice(later.index)
    : body.replace("</worksheet>", ignored + "</worksheet>");
}
function forceTextFormat(bytes) {
  const zip = XLSX.CFB.read(bytes, {
    type: "array",
  });
  const rel = (p) => p.replace(/^[^/]*\//, "");
  const dec = new TextDecoder("utf-8");
  const enc = new TextEncoder();
  const read = (i) => dec.decode(new Uint8Array(zip.FileIndex[i].content));
  const write = (i, text) => {
    const content = enc.encode(text);
    zip.FileIndex[i].content = content;
    zip.FileIndex[i].size = content.length;
  };
  const stylesIdx = zip.FullPaths.findIndex((p) => rel(p) === "xl/styles.xml");
  if (stylesIdx < 0) throw new Error("xlsx 안에 styles.xml이 없습니다.");
  const styles = textFormatStylesXml(read(stylesIdx));
  write(stylesIdx, styles.xml);
  let sheets = 0;
  zip.FullPaths.forEach((p, i) => {
    if (!/^xl\/worksheets\/[^/]+\.xml$/.test(rel(p))) return;
    write(i, textFormatSheetXml(read(i), styles.offset));
    sheets++;
  });
  if (sheets === 0) throw new Error("xlsx 안에 시트가 없습니다.");
  const out = XLSX.CFB.write(zip, {
    type: "array",
    fileType: "zip",
    compression: true,
  });
  return out instanceof Uint8Array ? out : new Uint8Array(out);
}
// never block a download: if anything goes wrong, hand out the file as it was (the parser still reads dates like 1월 1일 as 1-1)
function withTextCellFormat(bytes) {
  try {
    return forceTextFormat(bytes);
  } catch (e) {
    console.warn("엑셀 양식의 셀 서식을 텍스트로 바꾸지 못해 원래 양식으로 내려받습니다.", e);
    return bytes;
  }
}
function xlsxBytesWithTextCells(wb) {
  return withTextCellFormat(
    new Uint8Array(
      XLSX.write(wb, {
        type: "array",
        bookType: "xlsx",
      }),
    ),
  );
}
// ① 학기별 과목 목록 순서: 학점이 높은 과목부터 → 같은 학점이면 교과(군) 순(국어·수학·영어·사회·과학·체육·예술·생활교양,
// 그 밖의 교과(군)은 편제표에 처음 나온 순서) → 그래도 같으면 편제표에 적힌 순서. 학교마다 학점·교과(군) 이름이 달라도 그대로 동작합니다.
const COURSE_LIST_GROUP_ORDER = Object.keys(NATIONAL_REQUIRED_CREDITS);
function courseListGroupRank(group, seenGroups) {
  const g = toCurriculumGroup(group);
  const i = COURSE_LIST_GROUP_ORDER.findIndex(
    (k) => g === k || (g.startsWith(k) && /^[\s(（]/.test(g.slice(k.length))),
  );
  if (i !== -1) return i;
  const j = seenGroups.indexOf(g);
  return COURSE_LIST_GROUP_ORDER.length + (j === -1 ? seenGroups.length : j);
}
function sortCoursesForList(list, allCourses) {
  const seen = Array.from(new Set((allCourses || list).map((c) => toCurriculumGroup(c.group))));
  return list
    .map((c, i) => ({
      c,
      i,
      credit: Number(c.credit) || 0,
      rank: courseListGroupRank(c.group, seen),
    }))
    .sort((a, b) => b.credit - a.credit || a.rank - b.rank || a.i - b.i)
    .map((x) => x.c);
}
// ---------- downloadable base template ----------
function buildTemplateWorkbook() {
  const guide = [
    ["편제표 기본 양식 — 사용안내"],
    [],
    ["■ 이 파일에서 작성할 시트는 딱 하나, ‘편제표입력’ 시트입니다."],
    ["   나머지 시트는 참고용이며 지우거나 그대로 두어도 무방합니다."],
    [],
    ["■ ‘편제표입력’ 시트 컬럼 설명"],
    ["구분", "학교지정(전교생 자동이수) 또는 학생선택(수강신청으로 채워짐) 중 하나를 입력"],
    ["교과(군)", "국어/수학/영어/사회/과학/체육/예술/기술·가정/제2외국어·한문/교양 등 학교 편제에 맞게 자유롭게 입력"],
    ["과목구분", "공통/일반/진로/융합 등 (참고용, 계산에는 사용되지 않음)"],
    [
      "과목명",
      "실제 과목명. 수강신청 파일의 과목명과 최대한 동일하게 입력하면 정확도가 올라갑니다. (공백, Ⅰ/Ⅱ 표기 차이는 자동 인식)",
    ],
    ["기준학점", "2022 개정 교육과정 기준 학점"],
    ["운영학점", "학교가 실제 운영하는 학점. 비워두면 기준학점이 자동 적용됩니다."],
    ["개설학기", "1-1 / 1-2 / 2-1 / 2-2 / 3-1 / 3-2 (참고용)"],
    ["선택구분(택N)", "택4(16학점) 등 선택 그룹 메모 (참고용)"],
    [
      "필수이수학점",
      "★중요: 해당 교과(군)의 3년간 졸업 최소 이수학점. 같은 교과(군)에 속한 모든 행에 동일한 값을 반복 입력해주세요.",
    ],
    ["비고", "자유 메모"],
    [],
    ["■ 계산 방식"],
    ["· 학교지정 과목은 모든 학생이 자동으로 이수한 것으로 계산됩니다."],
    ["· 학생선택 과목은 수강신청 파일에서 실제로 신청(1로 표시)한 학생에게만 학점이 반영됩니다."],
    ["· 각 교과(군)의 이수학점 = 학교지정 학점 합 + 그 학생이 신청한 학생선택 학점 합"],
    [
      "· 위 값이 ‘필수이수학점’보다 낮은 교과(군)가 하나라도 있거나, 총학점이 최소 이수학점(웹사이트에서 설정)보다 낮으면 ‘확인필요’로 표시됩니다.",
    ],
    [],
    ["■ 작성이 끝나면 이 파일을 그대로 웹사이트 ①편제표 페이지에 업로드하세요."],
  ];
  const header = [
    "구분",
    "교과(군)",
    "과목구분",
    "과목명",
    "기준학점",
    "운영학점",
    "개설학기",
    "선택구분(택N)",
    "필수이수학점",
    "비고",
  ];
  const example = [
    ["학교지정", "국어", "공통", "공통국어1", 4, 4, "1-1", "필수(학교지정)", 8, ""],
    ["학교지정", "국어", "공통", "공통국어2", 4, 4, "1-2", "필수(학교지정)", 8, ""],
    ["학교지정", "수학", "공통", "공통수학1", 4, 4, "1-1", "필수(학교지정)", 8, ""],
    ["학생선택", "국어", "일반", "독서와 작문", 4, 4, "3-1", "택5(20학점)", 8, "예시 — 실제 편제에 맞게 수정하세요"],
  ];
  const wb = XLSX.utils.book_new();
  const wsGuide = XLSX.utils.aoa_to_sheet(guide);
  wsGuide["!cols"] = [
    {
      wch: 16,
    },
    {
      wch: 70,
    },
  ];
  XLSX.utils.book_append_sheet(wb, wsGuide, "사용안내");
  const wsInput = XLSX.utils.aoa_to_sheet([header, ...example]);
  wsInput["!cols"] = header.map(() => ({
    wch: 16,
  }));
  XLSX.utils.book_append_sheet(wb, wsInput, "편제표입력");
  return wb;
}
function buildTemplateXlsxBytes() {
  // ship the school's actual refined template (embedded as base64) instead of a generated one — all cells set to text
  const binary = atob(TEMPLATE_XLSX_BASE64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return withTextCellFormat(bytes);
}
function downloadTemplate() {
  downloadBlob(
    new Blob([buildTemplateXlsxBytes()], {
      type: "application/octet-stream",
    }),
    "편제표양식_입력용.xlsx",
  );
}
function buildClassMappingSampleBytes() {
  const rows = [
    ["1학년 반", "1학년 번호", "성명", "2학년 반", "2학년 번호"],
    [1, 5, "김민지", 2, 1],
    [3, 12, "김민지", 2, 15],
    [2, 3, "박서준", 1, 8],
  ];
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"] = rows[0].map(() => ({
    wch: 12,
  }));
  XLSX.utils.book_append_sheet(wb, ws, "학급편성변경");
  return xlsxBytesWithTextCells(wb);
}
function downloadClassMappingSample() {
  downloadBlob(
    new Blob([buildClassMappingSampleBytes()], {
      type: "application/octet-stream",
    }),
    "학급편성변경_샘플.xlsx",
  );
}
function formatSemester(code) {
  const m = /^(\d)-(\d)$/.exec(String(code || "").trim());
  if (!m) return code || "학기 미상";
  return `${m[1]}학년 ${m[2]}학기`;
}
function semesterSortKey(code) {
  const m = /^(\d)-(\d)$/.exec(String(code || "").trim());
  if (!m) return [99, 99];
  return [Number(m[1]), Number(m[2])];
}
function detectSemesterFromFilename(name) {
  if (!name) return null;
  const m1 = /([123])\s*학년\s*([12])\s*학기/.exec(name);
  if (m1) return `${m1[1]}-${m1[2]}`;
  const m2 = /(?<![0-9])([123])\s*[-_]\s*([12])(?![0-9])/.exec(name);
  if (m2) return `${m2[1]}-${m2[2]}`;
  return null;
}
const GROUP_PRIORITY = ["국어", "영어", "수학", "과학", "사회"];
function groupPriorityIndex(g) {
  const i = GROUP_PRIORITY.indexOf(g);
  return i === -1 ? 999 : i;
}
function parseRequiredN(label) {
  const m = /택\s*(\d+)/.exec(String(label || ""));
  return m ? Number(m[1]) : null;
}
// ---------- local save/load (optional password, client-side only — nothing leaves the computer) ----------
function bufToB64(buf) {
  let binary = "";
  const bytes = new Uint8Array(buf);
  for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
  return btoa(binary);
}
function b64ToBuf(b64) {
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}
async function deriveKey(password, saltBuf) {
  const enc = new TextEncoder();
  const baseKey = await crypto.subtle.importKey("raw", enc.encode(password), "PBKDF2", false, ["deriveKey"]);
  return crypto.subtle.deriveKey(
    {
      name: "PBKDF2",
      salt: saltBuf,
      iterations: 150000,
      hash: "SHA-256",
    },
    baseKey,
    {
      name: "AES-GCM",
      length: 256,
    },
    false,
    ["encrypt", "decrypt"],
  );
}
async function encryptText(text, password) {
  const enc = new TextEncoder();
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const key = await deriveKey(password, salt);
  const cipherBuf = await crypto.subtle.encrypt(
    {
      name: "AES-GCM",
      iv,
    },
    key,
    enc.encode(text),
  );
  return {
    encrypted: true,
    salt: bufToB64(salt),
    iv: bufToB64(iv),
    data: bufToB64(cipherBuf),
  };
}
async function decryptText(payload, password) {
  const salt = b64ToBuf(payload.salt);
  const iv = new Uint8Array(b64ToBuf(payload.iv));
  const key = await deriveKey(password, salt);
  const plainBuf = await crypto.subtle.decrypt(
    {
      name: "AES-GCM",
      iv,
    },
    key,
    b64ToBuf(payload.data),
  );
  return new TextDecoder().decode(plainBuf);
}
// 자동 보관용 잠금 도구. 파일 저장(encryptText)과 같은 형식으로 내보내서 decryptText로 열 수 있습니다.
async function makeKeeper(password) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const key = await deriveKey(password, salt);
  const saltB64 = bufToB64(salt);
  const enc = new TextEncoder();
  return {
    async encrypt(text) {
      const iv = crypto.getRandomValues(new Uint8Array(12));
      const cipherBuf = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, enc.encode(text));
      return { encrypted: true, salt: saltB64, iv: bufToB64(iv), data: bufToB64(cipherBuf) };
    },
  };
}
function toCsv(rows, headers) {
  const esc = (v) => {
    let s = v == null ? "" : String(v);
    // a text cell starting with = + @ (or - followed by text) would be run as a formula when opened in Excel
    if (typeof v === "string" && (/^[=+@\t\r]/.test(s) || /^-[^\d\s]/.test(s))) s = "'" + s;
    if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
  };
  const lines = [headers.map(esc).join(",")];
  rows.forEach((r) => lines.push(headers.map((h) => esc(r[h])).join(",")));
  return "\uFEFF" + lines.join("\n");
}
// ---------- small shared helpers ----------
function newId(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}
function entryKey(name, semester) {
  return `${normalizeName(name)}||${semester || ""}`;
}
function compareSemester(a, b) {
  const ka = semesterSortKey(a);
  const kb = semesterSortKey(b);
  return ka[0] - kb[0] || ka[1] - kb[1];
}
function toCreditNumber(v) {
  if (v == null) return 0;
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  const m = /\d+(\.\d+)?/.exec(String(v));
  return m ? Number(m[0]) : 0;
}
// Downloads a Blob as a file. The link is attached to the page and the URL is revoked a moment later —
// revoking synchronously or clicking a detached link silently cancels the download in some browsers.
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.style.display = "none";
  document.body.appendChild(a);
  a.click();
  setTimeout(() => {
    a.remove();
    URL.revokeObjectURL(url);
  }, 1500);
}
// ---------- tolerant reading of hand-typed 편제표 cells ----------
// '학교 지정', '필수(학교지정)', '학생 선택' 같은 표기 차이를 흡수합니다.
function normalizeGubun(v) {
  const s = String(v == null ? "" : v).replace(/\s+/g, "");
  if (/선택/.test(s)) return "학생선택";
  if (/지정|필수/.test(s)) return "학교지정";
  return String(v == null ? "" : v).trim();
}
// '2-1', '2학년 1학기', '2_1', and '2-1' that Excel silently turned into a date (2월 1일 → a date serial number).
function normalizeSemesterCode(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") {
    if (v > 59 && v < 2958465) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      const month = d.getUTCMonth() + 1;
      const day = d.getUTCDate();
      if (month >= 1 && month <= 3 && day >= 1 && day <= 2) return `${month}-${day}`;
    }
    return String(v);
  }
  const s = String(v).trim();
  let m = /^([123])\s*[-_.·/]\s*([12])(\s*학기)?$/.exec(s);
  if (m) return `${m[1]}-${m[2]}`;
  m = /([123])\s*학년\s*([12])\s*학기/.exec(s);
  if (m) return `${m[1]}-${m[2]}`;
  m = /^([123])\s*월\s*([12])\s*일$/.exec(s);
  if (m) return `${m[1]}-${m[2]}`;
  return s;
}
// 편제표 자체의 빈칸·오류를 미리 알려줍니다 (계산이 조용히 틀어지는 원인들).
function validateCurriculum(courses) {
  const issues = [];
  const push = (level, title, items) => {
    if (items.length > 0)
      issues.push({
        level,
        title,
        items,
      });
  };
  push(
    "error",
    "학점이 0이거나 비어 있는 과목 — 0학점으로 계산됩니다",
    courses.filter((c) => !(c.credit > 0)).map((c) => `${c.name} (${formatSemester(c.semester)})`),
  );
  push(
    "error",
    "개설학기를 알 수 없는 과목 — 1-1 ~ 3-2 형식으로 입력해주세요",
    courses.filter((c) => !SEMESTER_CODES.includes(c.semester)).map((c) => `${c.name} — ‘${c.semester || "빈칸"}’`),
  );
  push(
    "error",
    "구분이 ‘학교지정/학생선택’이 아닌 과목 — 전교생 자동 이수로 계산되지 않습니다",
    courses
      .filter((c) => c.gubun !== "학교지정" && c.gubun !== "학생선택")
      .map((c) => `${c.name} — ‘${c.gubun || "빈칸"}’`),
  );
  const noN = new Map();
  courses
    .filter((c) => c.gubun === "학생선택" && parseRequiredN(c.selectGroup) == null)
    .forEach((c) => {
      const k = `${c.semester}||${c.selectGroup}`;
      if (!noN.has(k))
        noN.set(k, {
          sem: c.semester,
          label: c.selectGroup,
          count: 0,
        });
      noN.get(k).count += 1;
    });
  push(
    "warn",
    "선택구분에 ‘택N’이 없는 학생선택 과목 — 택N 선택 개수 점검에서 빠집니다",
    Array.from(noN.values()).map((x) => `${formatSemester(x.sem)} ‘${x.label || "빈칸"}’ (${x.count}과목)`),
  );
  const dup = new Map();
  courses.forEach((c) => {
    const k = `${c.key}||${c.semester}`;
    dup.set(k, [...(dup.get(k) || []), c]);
  });
  push(
    "warn",
    "같은 학기에 같은 과목이 두 번 들어 있음",
    Array.from(dup.values())
      .filter((l) => l.length > 1)
      .map((l) => `${l[0].name} (${formatSemester(l[0].semester)}) × ${l.length}`),
  );
  return issues;
}
// ---------- registration pipeline (pure) ----------
// 고교학점제 파일은 과목마다 학기가 적혀 있고, 압핀 파일은 파일 하나가 한 학기라 파일 단위로 학기를 정합니다
// (직접 선택 > 파일 이름 > 과목 구성 추정). 학기를 끝내 알 수 없는 파일은 계산에 넣지 않습니다.
function computeRegFileInfo(f, courses) {
  const records = f.records || [];
  const needsFileSemester = records.some((r) => r.subjects.some((s) => !s.semester));
  const filenameSemester = detectSemesterFromFilename(f.name);
  const contentGuess = needsFileSemester
    ? guessSemesterFromSubjects(
        records.flatMap((r) => r.subjects.filter((s) => !s.semester).map((s) => s.name)),
        courses,
      )
    : null;
  let effectiveSemester = null;
  let semesterOrigin = null;
  if (needsFileSemester) {
    if (f.semesterMode && f.semesterMode !== "auto" && SEMESTER_CODES.includes(f.semesterMode)) {
      effectiveSemester = f.semesterMode;
      semesterOrigin = "manual";
    } else if (filenameSemester) {
      effectiveSemester = filenameSemester;
      semesterOrigin = "filename";
    } else if (contentGuess) {
      effectiveSemester = contentGuess.sem;
      semesterOrigin = "content";
    }
  }
  const semStudents = new Map();
  records.forEach((r) => {
    new Set(r.subjects.map((s) => s.semester || effectiveSemester).filter(Boolean)).forEach((sem) =>
      semStudents.set(sem, (semStudents.get(sem) || 0) + 1),
    );
  });
  const semesterCounts = Array.from(semStudents.entries())
    .map(([sem, students]) => ({
      sem,
      students,
    }))
    .sort((a, b) => compareSemester(a.sem, b.sem));
  return {
    needsFileSemester,
    filenameSemester,
    contentGuess,
    effectiveSemester,
    semesterOrigin,
    needsSemester: needsFileSemester && !effectiveSemester,
    semesterConflict: !!(
      needsFileSemester &&
      contentGuess &&
      effectiveSemester &&
      contentGuess.sem !== effectiveSemester
    ),
    semesterCounts,
    analysis: f.status === "ok" && courses.length > 0 ? analyzeRegFile(records, effectiveSemester, courses) : null,
  };
}
// which sources supply each semester, oldest first (the last one is the most recently added):
// 수강신청 파일(‘계산에 사용’을 끈 파일 제외)과 학생 기초조사 응답(응답이 있을 때)
function buildSemesterProviders(regFiles, infos, survey) {
  const entries = [];
  regFiles.forEach((f, i) => {
    if (f.status !== "ok" || f.disabled) return;
    const info = infos.get(f.id);
    if (!info || info.needsSemester) return;
    info.semesterCounts.forEach(({ sem }) =>
      entries.push({
        sem,
        id: f.id,
        at: Number(f.addedAt) || 0,
        order: i,
      }),
    );
  });
  if (survey && Array.isArray(survey.semesters))
    survey.semesters.forEach((sem) =>
      entries.push({
        sem,
        id: SURVEY_SOURCE_ID,
        at: Number(survey.addedAt) || 0,
        order: regFiles.length,
      }),
    );
  entries.sort((a, b) => a.at - b.at || a.order - b.order);
  const m = new Map();
  entries.forEach(({ sem, id }) => {
    if (!m.has(sem)) m.set(sem, []);
    if (!m.get(sem).includes(id)) m.get(sem).push(id);
  });
  return m;
}
// exactly one file per semester — never a union of two files' selections. default: the most recently uploaded file
function pickSourceBySem(providers, choice) {
  const m = new Map();
  providers.forEach((ids, sem) => {
    const picked = choice ? choice[sem] : null;
    m.set(sem, ids.includes(picked) ? picked : ids[ids.length - 1]);
  });
  return m;
}
function buildRawRecords(regFiles, infos, chosenSourceBySem) {
  const out = [];
  regFiles.forEach((f) => {
    if (f.status !== "ok" || f.disabled) return;
    const info = infos.get(f.id);
    if (!info || info.needsSemester) return;
    f.records.forEach((r, ri) => {
      var _a;
      const subjects = r.subjects
        .map((s) => ({
          ...s,
          semester: s.semester || info.effectiveSemester || null,
        }))
        .filter((s) => {
          var _a;
          return (
            !s.semester || ((_a = chosenSourceBySem.get(s.semester)) !== null && _a !== void 0 ? _a : f.id) === f.id
          );
        });
      if (r.subjects.length > 0 && subjects.length === 0) return; // this student's data for that semester comes from another file
      out.push({
        ...r,
        uid: `${f.id}:${ri}`,
        fileId: f.id,
        source: f.source,
        fileLabel: (_a = f.label) !== null && _a !== void 0 ? _a : f.name,
        fileSemester: info.effectiveSemester,
        subjects,
        semesters: Array.from(new Set(subjects.map((s) => s.semester).filter(Boolean))),
      });
    });
  });
  return out;
}
// names that appear more than once within a single block (one semester-table) → ambiguous
// (a name repeating across DIFFERENT blocks/files is expected — the same student across semesters)
function computeAmbiguousNames(regFiles) {
  const names = new Set();
  regFiles.forEach((f) => {
    if (f.status !== "ok" || f.disabled) return;
    const byBlock = new Map();
    (f.records || []).forEach((r) => {
      var _a;
      const bId = (_a = r.blockId) !== null && _a !== void 0 ? _a : "0";
      if (!byBlock.has(bId)) byBlock.set(bId, new Map());
      const c = byBlock.get(bId);
      c.set(r.name, (c.get(r.name) || 0) + 1);
    });
    byBlock.forEach((c) => c.forEach((cnt, name) => cnt > 1 && names.add(name)));
  });
  return names;
}
function buildCourseMap(courses) {
  const m = new Map();
  courses.forEach((c) => {
    if (!m.has(c.key)) m.set(c.key, []);
    m.get(c.key).push(c);
  });
  return m;
}
function buildSemesterPoolDefs(courses) {
  const bySemPool = new Map();
  courses
    .filter((c) => c.gubun === "학생선택")
    .forEach((c) => {
      const key = `${c.semester}||${c.selectGroup}`;
      if (!bySemPool.has(key))
        bySemPool.set(key, {
          sem: c.semester,
          label: c.selectGroup,
          requiredN: parseRequiredN(c.selectGroup),
          courses: [],
        });
      bySemPool.get(key).courses.push(c);
    });
  return Array.from(bySemPool.values())
    .map((p) => ({
      ...p,
      courses: [...p.courses].sort((a, b) => groupPriorityIndex(a.group) - groupPriorityIndex(b.group)),
    }))
    .sort((a, b) => compareSemester(a.sem, b.sem) || String(a.label).localeCompare(String(b.label)));
}
function manualStudentIdFields(ms) {
  const grade = Number(ms.grade) || null;
  const classNum = Number(ms.classNum) || null;
  const number = Number(ms.number) || null;
  return {
    grade,
    classNum,
    number,
  };
}
function buildStudentGroups(rawRecords, groupKeyFor, manualStudents) {
  const map = new Map();
  rawRecords.forEach((rec) => {
    const key = groupKeyFor(rec);
    if (!map.has(key))
      map.set(key, {
        key,
        name: rec.name,
        records: [],
        subjectEntries: new Map(),
      });
    const g = map.get(key);
    g.records.push(rec);
    rec.subjects.forEach((s) => {
      const sem = s.semester || rec.fileSemester || null;
      const k = entryKey(s.name, sem);
      if (!g.subjectEntries.has(k))
        g.subjectEntries.set(k, {
          name: s.name,
          semester: sem,
        });
    });
  });
  // 수강신청 파일에 없는 학생(나중에 온 전입생 등)을 ⑤에서 직접 추가한 경우
  (manualStudents || []).forEach((ms) => {
    const key = "manual:" + ms.id;
    if (map.has(key)) return;
    const idFields = manualStudentIdFields(ms);
    const full = idFields.grade && idFields.classNum && idFields.number;
    const idKey = full
      ? `${idFields.grade}${String(idFields.classNum).padStart(2, "0")}${String(idFields.number).padStart(2, "0")}`
      : "-";
    map.set(key, {
      key,
      name: ms.name,
      manual: true,
      subjectEntries: new Map(),
      records: [
        {
          uid: key,
          name: ms.name,
          idFields: full ? idFields : null,
          idKey,
          idText: full ? `${idFields.grade}학년 ${idFields.classNum}반 ${idFields.number}번` : "학번 미입력",
          subjects: [],
          semesters: [],
          fileLabel: "직접 추가",
          source: "manual",
        },
      ],
    });
  });
  return Array.from(map.values());
}
// the most recent grade present in the data (2 when a 고교학점제 file carries 2학년 학번, 1 when only 1학년 때 받은 압핀
// files were uploaded) — a student with no record at that grade most likely left the school.
function computeLatestGrade(studentGroups) {
  let mx = 1;
  studentGroups.forEach((g) =>
    g.records.forEach((r) => {
      if (r.idFields && r.idFields.grade) mx = Math.max(mx, r.idFields.grade);
    }),
  );
  return mx;
}
// ---------- student result computation (pure) ----------
const UNCLASSIFIED_LABEL = "확인 필요 (반 정보 없음)";
// 전입생: firstSem(본교에서 처음 다닌 학기)보다 앞 학기는 본교 편제표 대신 전적교 이수 과목으로 계산합니다.
function isAtThisSchool(semester, firstSem) {
  if (!firstSem) return true;
  const i = SEMESTER_CODES.indexOf(semester);
  const f = SEMESTER_CODES.indexOf(firstSem);
  if (i < 0 || f < 0) return true;
  return i >= f;
}
function semestersBefore(firstSem) {
  const f = SEMESTER_CODES.indexOf(firstSem);
  return f <= 0 ? [] : SEMESTER_CODES.slice(0, f);
}
function computeStudentResult(g, ctx) {
  var _a;
  const {
    courses,
    groups,
    subjectMaster,
    courseMap,
    poolDefs,
    reqOverride,
    minCredit,
    latestGradeInData,
    coveredSemesters,
  } = ctx;
  const change = ctx.change || null;
  const transfer = ctx.transfer || null;
  const firstSem = transfer && SEMESTER_CODES.includes(transfer.firstSem) ? transfer.firstSem : null;
  const here = (sem) => isAtThisSchool(sem, firstSem);
  const groupCredit = {};
  groups.forEach((gr) => (groupCredit[gr] = 0));
  const addCredit = (group, credit) => {
    const gg = toCurriculumGroup(group) || "기타";
    groupCredit[gg] = (groupCredit[gg] || 0) + (Number(credit) || 0);
  };
  // 1) 학교지정 과목 — 본교에 다닌 학기만
  courses.forEach((c) => {
    if (c.gubun === "학교지정" && here(c.semester)) addCredit(c.group, c.credit);
  });
  // 2) 수강신청 기록 → 상담 후 변경(취소/추가) 적용
  const removedSet = new Set(change && Array.isArray(change.removed) ? change.removed : []);
  const registered = [];
  const removedEntries = [];
  const excludedBeforeTransfer = [];
  const entries = [];
  const seen = new Set();
  Array.from(g.subjectEntries.values()).forEach((e) => {
    if (!here(e.semester)) {
      excludedBeforeTransfer.push({
        name: e.name,
        semester: e.semester,
      });
      return;
    }
    const k = entryKey(e.name, e.semester);
    if (seen.has(k)) return;
    seen.add(k);
    registered.push({
      name: e.name,
      semester: e.semester,
    });
    if (removedSet.has(k)) {
      removedEntries.push({
        name: e.name,
        semester: e.semester,
      });
      return;
    }
    entries.push({
      name: e.name,
      semester: e.semester,
      origin: "registration",
    });
  });
  const addedEntries = [];
  (change && Array.isArray(change.added) ? change.added : []).forEach((a) => {
    if (!a || !a.name) return;
    const k = entryKey(a.name, a.semester);
    if (seen.has(k) || !here(a.semester)) return;
    seen.add(k);
    entries.push({
      name: a.name,
      semester: a.semester || null,
      origin: "added",
    });
    addedEntries.push({
      name: a.name,
      semester: a.semester || null,
    });
  });
  const unmatched = [];
  const notInCurriculum = [];
  const semesterMismatch = [];
  const taken = [];
  const withSem = (n, sem) => (sem ? `${n} (${formatSemester(sem)})` : n);
  entries.forEach(({ name: subName, semester: takenSem, origin }) => {
    const key = normalizeName(subName);
    const list = courseMap.get(key);
    const { course: c, mismatch } = resolveTakenCourse(list, takenSem);
    if (c) {
      const offered = Array.from(new Set(list.map((x) => x.semester).filter(Boolean)));
      if (mismatch)
        semesterMismatch.push({
          name: subName,
          semester: takenSem,
          offered,
        });
      if (c.gubun === "학교지정" && here(c.semester)) {
        // already counted as 학교지정 — a registration file listing it again must not add the credit twice
        taken.push({
          name: subName,
          group: c.group,
          credit: c.credit,
          semester: takenSem || c.semester,
          selectGroup: null,
          origin,
          schoolDesignated: true,
        });
        return;
      }
      addCredit(c.group, c.credit);
      taken.push({
        name: subName,
        group: toCurriculumGroup(c.group),
        credit: c.credit,
        semester: mismatch ? takenSem : c.semester || takenSem || null,
        selectGroup: mismatch ? null : c.selectGroup || null,
        origin,
      });
      return;
    }
    const master = subjectMaster && typeof subjectMaster.get === "function" ? subjectMaster.get(key) : null;
    if (master) {
      addCredit(master.group, master.credit);
      taken.push({
        name: subName,
        group: toCurriculumGroup(master.group),
        credit: master.credit,
        semester: takenSem || null,
        selectGroup: null,
        origin,
      });
      notInCurriculum.push(withSem(subName, takenSem));
      return;
    }
    unmatched.push(withSem(subName, takenSem));
  });
  // 3) 전입 이전 학기 — 전적교 이수 과목
  const priorCourses = [];
  const priorIgnored = [];
  if (firstSem) {
    (Array.isArray(transfer.priorCourses) ? transfer.priorCourses : []).forEach((pc) => {
      if (!pc || !String(pc.name || "").trim()) return;
      const row = {
        id: pc.id,
        semester: pc.semester,
        group: toCurriculumGroup(pc.group),
        name: String(pc.name).trim(),
        credit: Number(pc.credit) || 0,
      };
      const before = SEMESTER_CODES.includes(pc.semester) && !isAtThisSchool(pc.semester, firstSem);
      if (!before) {
        priorIgnored.push(row);
        return;
      }
      addCredit(row.group, row.credit);
      priorCourses.push(row);
    });
    priorCourses.sort((a, b) => compareSemester(a.semester, b.semester));
  }
  // 4) 공동교육과정·온라인학교 등 추가 이수
  const extraEntries = [];
  (ctx.extras || []).forEach((x) => {
    if (!x || !String(x.name || "").trim()) return;
    const row = {
      id: x.id,
      kind: x.kind || "기타",
      name: String(x.name).trim(),
      group: toCurriculumGroup(x.group),
      credit: Number(x.credit) || 0,
      semester: x.semester || null,
      memo: x.memo || "",
    };
    addCredit(row.group, row.credit);
    extraEntries.push(row);
  });
  // 5) 택N 선택 현황 (본교 학기만)
  const pools = poolDefs
    .filter((def) => here(def.sem))
    .map((def) => {
      const takenCourses = def.courses.filter((c) =>
        entries.some((t) => normalizeName(t.name) === c.key && (!t.semester || t.semester === def.sem)),
      );
      const n = def.requiredN;
      return {
        sem: def.sem,
        label: def.label,
        requiredN: n,
        takenCount: takenCourses.length,
        takenCourses,
        deficient: n != null && takenCourses.length < n,
        over: n != null && takenCourses.length > n,
      };
    });
  // 6) 같은 과목을 같은 학년 1·2학기에 모두 선택
  const duplicateSelections = [];
  {
    const byNameGrade = new Map();
    entries.forEach(({ name: subName, semester: takenSem }) => {
      const m = /^(\d)-(\d)$/.exec(takenSem || "");
      if (!m) return;
      const gradeKey = `${normalizeName(subName)}||${m[1]}`;
      if (!byNameGrade.has(gradeKey))
        byNameGrade.set(gradeKey, {
          name: subName,
          grade: m[1],
          terms: new Set(),
        });
      byNameGrade.get(gradeKey).terms.add(m[2]);
    });
    byNameGrade.forEach((v) => {
      if (v.terms.size > 1)
        duplicateSelections.push({
          name: v.name,
          semesters: Array.from(v.terms)
            .sort()
            .map((t) => `${v.grade}-${t}`),
        });
    });
  }
  // 7) 신청 기록이 아예 없는 학기 (그 학기 파일은 올라왔는데 이 학생만 없음)
  const poolSems = Array.from(new Set(pools.map((p) => p.sem)));
  const entrySems = new Set(entries.map((e) => e.semester).filter(Boolean));
  const registeredSems = new Set(registered.map((e) => e.semester).filter(Boolean));
  const missingSemesters = poolSems
    .filter(
      (sem) => (!coveredSemesters || coveredSemesters.has(sem)) && !entrySems.has(sem) && !registeredSems.has(sem),
    )
    .sort(compareSemester);
  const round = (v) => Math.round(v * 100) / 100;
  Object.keys(groupCredit).forEach((k) => (groupCredit[k] = round(groupCredit[k])));
  const total = round(Object.values(groupCredit).reduce((a, b) => a + b, 0));
  const reqOf = (gr) => Number(reqOverride[gr]) || 0;
  const shortGroups = groups.filter((gr) => (groupCredit[gr] || 0) < reqOf(gr));
  const otherGroups = Object.keys(groupCredit)
    .filter((gr) => !groups.includes(gr) && groupCredit[gr] > 0)
    .map((gr) => ({
      group: gr,
      credit: groupCredit[gr],
    }));
  const pass = shortGroups.length === 0 && total >= minCredit && duplicateSelections.length === 0;
  const reasons = [];
  shortGroups.forEach((gr) =>
    reasons.push({
      type: "short",
      blocking: true,
      group: gr,
      text: `${gr} ${round(reqOf(gr) - (groupCredit[gr] || 0))}학점 부족`,
    }),
  );
  if (total < minCredit)
    reasons.push({
      type: "total",
      blocking: true,
      text: `총 ${total}/${minCredit}학점`,
    });
  duplicateSelections.forEach((d) =>
    reasons.push({
      type: "dup",
      blocking: true,
      text: `${d.name} 두 학기 중복 선택`,
    }),
  );
  missingSemesters.forEach((sem) =>
    reasons.push({
      type: "missing",
      blocking: false,
      text: `${formatSemester(sem)} 신청 기록 없음`,
    }),
  );
  pools
    .filter((p) => !missingSemesters.includes(p.sem))
    .forEach((p) => {
      if (p.deficient)
        reasons.push({
          type: "pool",
          blocking: false,
          text: `${formatSemester(p.sem)} ${p.label} ${p.takenCount}/${p.requiredN}`,
        });
      if (p.over)
        reasons.push({
          type: "pool",
          blocking: false,
          text: `${formatSemester(p.sem)} ${p.label} ${p.takenCount}/${p.requiredN} (초과)`,
        });
    });
  if (semesterMismatch.length)
    reasons.push({
      type: "mismatch",
      blocking: false,
      text: `편제표와 학기 다름 ${semesterMismatch.length}과목`,
    });
  if (unmatched.length)
    reasons.push({
      type: "unmatched",
      blocking: false,
      text: `인식 못한 과목 ${unmatched.length}`,
    });
  // 학급·학번 표시
  const records = g.records || [];
  const isUnclassified = !records.some((r) => r.idFields && r.idFields.grade >= latestGradeInData);
  const repRecord =
    records.find(
      (r) =>
        r.idFields &&
        r.idFields.grade &&
        r.idFields.classNum &&
        (r.semesters || []).some((s) => s === "3-1" || s === "3-2"),
    ) ||
    records.find((r) => r.idFields && r.idFields.grade >= 2 && r.idFields.classNum) ||
    records.find((r) => r.idFields && r.idFields.grade && r.idFields.classNum) ||
    records[0];
  const repOk = !!(repRecord && repRecord.idFields && repRecord.idFields.grade && repRecord.idFields.classNum);
  const classLabel =
    !isUnclassified && repOk ? `${repRecord.idFields.grade}학년 ${repRecord.idFields.classNum}반` : UNCLASSIFIED_LABEL;
  const classSortKey = !isUnclassified && repOk ? repRecord.idFields.grade * 100 + repRecord.idFields.classNum : 99999;
  const changeMemo = change && change.memo ? String(change.memo) : "";
  // 수강신청 시스템 반영 표시: 표시할 때의 변경 내역과 지금 내역이 같을 때만 ‘반영함’
  const hasEdits = addedEntries.length + removedEntries.length > 0;
  const appliedMark = change && change.applied && typeof change.applied === "object" ? change.applied : null;
  const appliedValid = !!(appliedMark && appliedMark.sig === changeSignature(change));
  const applyState = !hasEdits ? "none" : appliedValid ? "applied" : appliedMark ? "stale" : "pending";
  return {
    key: g.key,
    name: g.name,
    isManual: !!g.manual,
    idLabel: records.map((r) => `${r.fileLabel}: ${r.idText}`).join(" / "),
    studentId: String(repRecord ? ((_a = repRecord.idKey) !== null && _a !== void 0 ? _a : "-") : "-"),
    classLabel,
    classSortKey,
    isUnclassified: classLabel === UNCLASSIFIED_LABEL,
    groupCredit,
    otherGroups,
    total,
    shortGroups,
    pass,
    reasons,
    taken,
    entries,
    registered,
    unmatched,
    notInCurriculum,
    semesterMismatch,
    pools,
    duplicateSelections,
    missingSemesters,
    change: {
      added: addedEntries,
      removed: removedEntries,
      memo: changeMemo,
      updatedAt: (change && change.updatedAt) || null,
      applyState,
      appliedAt: appliedValid ? appliedMark.at || null : null,
    },
    hasChanges: addedEntries.length + removedEntries.length > 0 || !!changeMemo,
    isTransfer: !!firstSem,
    firstSem,
    prevSchool: transfer ? String(transfer.prevSchool || "") : "",
    transferMemo: transfer ? String(transfer.memo || "") : "",
    priorCourses,
    priorIgnored,
    excludedBeforeTransfer,
    extraEntries,
  };
}
function computeAllResults(studentGroups, ctx, adj) {
  const extrasByStudent = new Map();
  (adj.extraCourses || []).forEach((x) =>
    (x.studentKeys || []).forEach((k) => {
      if (!extrasByStudent.has(k)) extrasByStudent.set(k, []);
      extrasByStudent.get(k).push(x);
    }),
  );
  const changes = adj.courseChanges || {};
  const transfers = adj.transferInfo || {};
  return studentGroups
    .map((g) =>
      computeStudentResult(g, {
        ...ctx,
        change: changes[g.key],
        transfer: transfers[g.key],
        extras: extrasByStudent.get(g.key) || [],
      }),
    )
    .sort((a, b) => a.name.localeCompare(b.name, "ko"));
}
// subject-name → {name, group, credit} lookup for the manual-entry forms (school's own 편제표 wins over the national list)
function buildSubjectLookup(courses) {
  const m = new Map();
  NATIONAL_SUBJECT_MASTER.forEach((s) =>
    m.set(normalizeName(s.name), {
      name: s.name,
      group: toCurriculumGroup(s.group),
      credit: s.credit,
    }),
  );
  (courses || []).forEach((c) =>
    m.set(c.key || normalizeName(c.name), {
      name: c.name,
      group: toCurriculumGroup(c.group),
      credit: c.credit,
    }),
  );
  return m;
}
// ---------- 나이스 ‘학생부 항목별 조회 → 교과학습발달상황’ 엑셀: 전입생 전적교 이수 과목 ----------
// 한 파일 = 한 반. 열: 번호 · 성명 · 학년 · 학기 · 교과 · 과목 · 학점(학점수) · 원점수 … (성적 열은 읽지 않습니다).
// 번호·성명·학년·학기 칸은 학생·학기마다 병합되어 첫 줄에만 값이 있으므로 아래로 채워 읽고, ‘이수학점 합계’ 줄로 학생을 끊습니다.
// 쪽이 바뀔 때 머리글이 되풀이되거나 제목·쪽번호 줄이 끼어 있어도 건너뜁니다.
const NEIS_HEADER_LABELS = {
  number: ["번호"],
  name: ["성명", "이름"],
  grade: ["학년"],
  term: ["학기"],
  group: ["교과"],
  subject: ["과목", "과목명"],
  credit: ["학점", "학점수", "단위수", "이수단위", "단위"],
};
function neisText(v) {
  return v == null ? "" : String(v).replace(/\s+/g, "");
}
function neisHeaderCols(row) {
  const cells = (row || []).map(neisText);
  const cols = {};
  Object.entries(NEIS_HEADER_LABELS).forEach(([key, labels]) => {
    const i = cells.findIndex((c) => labels.includes(c));
    if (i !== -1) cols[key] = i;
  });
  return cols.name != null && cols.subject != null && cols.credit != null && cols.grade != null && cols.term != null
    ? cols
    : null;
}
function parseNeisGradesWorkbook(wb) {
  let classInfo = null;
  let headerSeen = false;
  const students = [];
  const warnings = [];
  wb.SheetNames.forEach((sn) => {
    const rows = sheetToMatrix(wb, sn) || [];
    let cols = null;
    let cur = null;
    let grade = null;
    let term = null;
    let group = "";
    rows.forEach((r) => {
      var _a, _b;
      if (!r) return;
      if (!classInfo) {
        for (const c of r) {
          const m = /^([123])학년(\d{1,2})반$/.exec(neisText(c));
          if (m) {
            classInfo = {
              grade: Number(m[1]),
              classNum: Number(m[2]),
            };
            break;
          }
        }
      }
      const header = neisHeaderCols(r);
      if (header) {
        cols = header;
        headerSeen = true;
        return; // 첫 머리글 또는 쪽마다 되풀이되는 머리글
      }
      if (!cols) return;
      if (r.some((c) => neisText(c).includes("이수학점"))) {
        if (cur) {
          const nums = r.map((c) => Number(neisText(c))).filter((n) => Number.isFinite(n) && n > 0);
          cur.total = nums.length ? nums[nums.length - 1] : null;
        }
        cur = null;
        grade = null;
        term = null;
        group = "";
        return;
      }
      const name = String((_a = r[cols.name]) !== null && _a !== void 0 ? _a : "")
        .replace(/\s+/g, " ")
        .trim();
      const numRaw = cols.number != null ? neisText(r[cols.number]) : "";
      const number = /^\d+$/.test(numRaw) ? Number(numRaw) : null;
      if (name && !/^[\d./:\s]+$/.test(name)) {
        if (!cur || cur.name !== name || (number != null && cur.number != null && number !== cur.number)) {
          cur = {
            number,
            name,
            courses: [],
            total: null,
          };
          students.push(cur);
          grade = null;
          term = null;
          group = "";
        }
      }
      if (!cur) return;
      const g = neisText(r[cols.grade]);
      const t = neisText(r[cols.term]);
      if (/^[123]$/.test(g)) grade = Number(g);
      if (/^[12]$/.test(t)) term = Number(t);
      const grp = cols.group != null ? neisText(r[cols.group]) : "";
      if (grp) group = grp;
      const subject = String((_b = r[cols.subject]) !== null && _b !== void 0 ? _b : "")
        .replace(/\s+/g, " ")
        .trim();
      const creditRaw = neisText(r[cols.credit]);
      if (!subject || !/^\d+(\.\d+)?$/.test(creditRaw)) return;
      if (!grade || !term) {
        warnings.push(`${cur.name}: ‘${subject}’의 학년·학기를 알 수 없어 뺐습니다.`);
        return;
      }
      cur.courses.push({
        semester: `${grade}-${term}`,
        group,
        name: subject,
        credit: Number(creditRaw),
      });
    });
  });
  if (!headerSeen) {
    throw new Error(
      "나이스 교과학습발달상황 엑셀이 아닙니다. (번호·성명·학년·학기·교과·과목·학점 머리글을 찾지 못했습니다)",
    );
  }
  const withCourses = students.filter((s) => s.courses.length > 0);
  if (withCourses.length === 0)
    throw new Error(
      "머리글은 찾았지만 학생 과목을 한 줄도 읽지 못했습니다. 나이스에서 내려받은 원본 엑셀인지 확인해주세요.",
    );
  withCourses.forEach((s) => {
    const sum = s.courses.reduce((a, c) => a + c.credit, 0);
    if (s.total != null && Math.abs(sum - s.total) > 0.001)
      warnings.push(`${s.name}: 과목 학점 합 ${sum}이 이수학점 합계 ${s.total}와 다릅니다.`);
  });
  return {
    classInfo,
    students: withCourses,
    warnings,
  };
}
// 나이스 교과 이름 → 이 학교 교과(군). ‘사회(역사/도덕포함)’ → 사회, ‘기술・가정/제2외국어/한문/교양’ → 생활교양,
// 그래도 모르면 과목명으로 2022 개정 과목 목록·편제표에서 찾고, 끝내 모르면 빈칸(가져온 뒤 직접 고름).
function neisGroupToCurriculum(raw, groups, subjectHit) {
  const list = groups || [];
  const g0 = toCurriculumGroup(raw);
  if (list.includes(g0)) return g0;
  const g1 = toCurriculumGroup(String(raw || "").replace(/[(（].*$/, ""));
  if (list.includes(g1)) return g1;
  if (subjectHit && list.includes(subjectHit.group)) return subjectHit.group;
  return "";
}
// 본교에서 처음 다닌 학기 이전 학기의 과목만 전적교 이수 과목 줄로 바꿉니다.
function neisPriorRows(student, firstSem, groups, subjectLookup) {
  const before = semestersBefore(firstSem);
  const rows = [];
  const skipped = [];
  (student && student.courses ? student.courses : []).forEach((c) => {
    if (!before.includes(c.semester)) {
      skipped.push(c);
      return;
    }
    const hit = subjectLookup ? subjectLookup.get(normalizeName(c.name)) : null;
    rows.push({
      id: newId("pc"),
      semester: c.semester,
      group: neisGroupToCurriculum(c.group, groups, hit),
      name: c.name,
      credit: c.credit,
    });
  });
  const semesters = Array.from(new Set(rows.map((x) => x.semester))).sort(compareSemester);
  return {
    rows,
    skipped,
    semesters,
    credit: rows.reduce((a, x) => a + x.credit, 0),
  };
}
// 이름이 같은 나이스 학생을 찾습니다. 파일의 반과 번호까지 지금 학번과 맞는 학생이 있으면 그 학생만 돌려줍니다.
function findNeisCandidates(neisFiles, result) {
  const same = (a, b) => String(a || "").replace(/\s+/g, "") === String(b || "").replace(/\s+/g, "");
  const id = decodeIdFromDigits(result && result.studentId);
  const out = [];
  (neisFiles || []).forEach((f) => {
    if (f.status !== "ok") return;
    (f.students || []).forEach((s, index) => {
      if (!same(s.name, result && result.name)) return;
      const exact = !!(
        id &&
        f.classInfo &&
        f.classInfo.grade === id.grade &&
        f.classInfo.classNum === id.classNum &&
        s.number === id.number
      );
      out.push({
        fileId: f.id,
        fileName: f.name,
        classInfo: f.classInfo,
        index,
        student: s,
        exact,
      });
    });
  });
  const exact = out.filter((c) => c.exact);
  return exact.length ? exact : out;
}
function neisPlaceText(c) {
  const cls = c.classInfo ? `${c.classInfo.grade}학년 ${c.classInfo.classNum}반` : "반 정보 없음";
  return c.student.number != null ? `${cls} ${c.student.number}번` : cls;
}
// ---------- save / load (진행 상황 파일) ----------
const SNAPSHOT_VERSION = 5; // 3: 학생 기초조사 · 4: 수강신청 파일과 기초조사를 학기별로 함께 쓰기(파일별 ‘계산에 사용’, 1학년 선택과목 확인, 결번) · 5: ④ 변경의 수강신청 시스템 반영 표시
const EXTRA_KINDS = ["공동교육과정", "온라인학교", "학교 밖 교육", "대학 연계", "기타", "이전 버전 입력"];
function nameFromStudentKey(key) {
  const s = String(key || "");
  if (s.startsWith("name:")) return s.slice(5);
  if (s.startsWith("id:")) return s.slice(3).split("||")[0];
  if (s.startsWith("mapped:")) return s.slice(7).split("||")[0];
  return "";
}
function normalizeBuilderData(bd) {
  const out = emptyBuilderData();
  if (!bd || typeof bd !== "object") return out;
  SEMESTER_CODES.forEach((sem) => {
    const s = bd[sem];
    if (!s || typeof s !== "object") return;
    out[sem] = {
      mandatory: (Array.isArray(s.mandatory) ? s.mandatory : [])
        .filter((c) => c && c.name)
        .map((c) => ({
          ...c,
          credit: Number(c.credit) || 0,
        })),
      pools: (Array.isArray(s.pools) ? s.pools : [])
        .filter((p) => p && typeof p === "object")
        .map((p, i) => ({
          id: p.id || `pool-${sem}-${i}`,
          requiredN: Number(p.requiredN) || 0,
          creditPerCourse: Number(p.creditPerCourse) || 0,
          courses: (Array.isArray(p.courses) ? p.courses : []).filter((c) => c && c.name),
        })),
    };
  });
  return out;
}
// Validates and upgrades a parsed 진행 상황 file (any version this program has written) into the current state
// shape. Throws an Error with a message a teacher can act on when the file isn't one of ours.
function normalizeSnapshot(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw))
    throw new Error("진행 상황 파일의 내용이 올바르지 않습니다.");
  const looksLikeOurs =
    raw.version != null || ["curriculumFile", "regFiles", "reqOverride", "builderData"].some((k) => k in raw);
  if (!looksLikeOurs) throw new Error("이 프로그램에서 ‘진행 상황 저장’으로 만든 파일이 아닙니다.");
  if (Number(raw.version) > SNAPSHOT_VERSION)
    throw new Error("더 새로운 버전의 프로그램에서 저장한 파일입니다. 최신 프로그램에서 열어주세요.");
  const arr = (v) => (Array.isArray(v) ? v : []);
  const obj = (v) => (v && typeof v === "object" && !Array.isArray(v) ? v : {});
  const str = (v) => (v == null ? "" : String(v));
  const creditOrBlank = (v) => (v === "" || v == null ? "" : Number(v) || 0);
  let curriculumFile = null;
  let curriculumReq = null;
  const cf = raw.curriculumFile;
  if (cf && typeof cf === "object" && Array.isArray(cf.courses)) {
    const courses = cf.courses
      .filter((c) => c && c.name)
      .map((c) => ({
        ...c,
        name: str(c.name),
        key: normalizeName(c.name),
        group: normalizeGroupName(c.group),
        gubun: normalizeGubun(c.gubun),
        credit: Number(c.credit) || 0,
        semester: normalizeSemesterCode(c.semester),
        selectGroup: str(c.selectGroup),
      }));
    const savedGroups =
      arr(cf.groups).length > 0
        ? arr(cf.groups).map(normalizeGroupName)
        : Array.from(new Set(courses.map((c) => c.group)));
    // Re-derive everything that depends on the matching rules instead of trusting what an older version saved:
    // version 1 stored keys like '기술⋅가정' with an un-normalized dot, so those registrations would silently stop
    // matching after loading, and it could leave 생활·교양 groups un-merged (no 16학점 기준).
    const merged = mergeLifestyleGroups(courses, obj(raw.reqOverride), savedGroups);
    curriculumFile = {
      name: str(cf.name) || "불러온 편제표",
      origin: cf.origin === "builder" ? "builder" : "upload",
      courses: merged.courses,
      groups: merged.groups,
      subjectMaster: new Map(
        arr(cf.subjectMasterEntries)
          .filter((e) => Array.isArray(e) && e.length === 2 && e[1] && typeof e[1] === "object")
          .map(([k, v]) => [
            normalizeName(k),
            {
              ...v,
              group: toCurriculumGroup(v.group),
              credit: Number(v.credit) || 0,
            },
          ]),
      ),
    };
    curriculumReq = merged.req;
  }
  const regFiles = arr(raw.regFiles)
    .filter((f) => f && typeof f === "object")
    .map((f, fi) => {
      const records = arr(f.records)
        .filter((r) => r && typeof r === "object" && r.name)
        .map((r) => ({
          ...r,
          name: str(r.name),
          subjects: arr(r.subjects)
            .filter((s) => s && s.name)
            .map((s) => ({
              ...s,
              name: str(s.name),
              semester: s.semester ? str(s.semester) : null,
            })),
          semesters: arr(r.semesters),
          ...(Array.isArray(r.pastIds)
            ? {
                pastIds: r.pastIds
                  .filter((p) => p && typeof p === "object")
                  .map((p) => ({
                    grade: Number(p.grade) || null,
                    classNum: Number(p.classNum) || null,
                    number: Number(p.number) || null,
                  })),
              }
            : {}),
        }));
      const inferred = records[0] && String(records[0].blockId || "").startsWith("B") ? "apin" : "hakjeom";
      return {
        ...f,
        id: f.id != null ? str(f.id) : String(fi),
        // older saves: keeps `${fileIndex}:${recordIndex}` uids so saved merges still apply
        source: f.source === "apin" || f.source === "hakjeom" ? f.source : inferred,
        name: str(f.name),
        label: f.label != null ? str(f.label) : str(f.name),
        semesterMode: f.semesterMode || "auto",
        status: ["ok", "error", "wrongZone"].includes(f.status) ? f.status : "ok",
        records,
        count: records.length,
        disabled: f.disabled === true,
        addedAt: Number(f.addedAt) || 0,
      };
    });
  const courseChanges = {};
  Object.entries(obj(raw.courseChanges)).forEach(([k, v]) => {
    if (!v || typeof v !== "object") return;
    courseChanges[k] = {
      name: str(v.name) || nameFromStudentKey(k),
      removed: arr(v.removed).map(str),
      added: arr(v.added)
        .filter((a) => a && a.name)
        .map((a) => ({
          name: str(a.name),
          semester: a.semester ? str(a.semester) : null,
        })),
      memo: str(v.memo),
      updatedAt: v.updatedAt || null,
      applied:
        v.applied && typeof v.applied === "object" && typeof v.applied.sig === "string"
          ? {
              at: str(v.applied.at),
              sig: v.applied.sig,
            }
          : null,
    };
  });
  const transferInfo = {};
  Object.entries(obj(raw.transferInfo)).forEach(([k, v]) => {
    if (!v || typeof v !== "object") return;
    transferInfo[k] = {
      name: str(v.name) || nameFromStudentKey(k),
      firstSem: SEMESTER_CODES.includes(v.firstSem) ? v.firstSem : "2-1",
      prevSchool: str(v.prevSchool),
      memo: str(v.memo),
      priorCourses: arr(v.priorCourses)
        .filter((p) => p && typeof p === "object")
        .map((p) => ({
          id: p.id ? str(p.id) : newId("pc"),
          semester: str(p.semester),
          group: str(p.group),
          name: str(p.name),
          credit: creditOrBlank(p.credit),
        })),
    };
  });
  const extraCourses = arr(raw.extraCourses)
    .filter((x) => x && typeof x === "object")
    .map((x) => ({
      id: x.id ? str(x.id) : newId("ex"),
      kind: EXTRA_KINDS.includes(x.kind) ? x.kind : "기타",
      name: str(x.name),
      group: str(x.group),
      credit: creditOrBlank(x.credit),
      semester: str(x.semester),
      memo: str(x.memo),
      studentKeys: Array.from(new Set(arr(x.studentKeys).map(str))),
      studentNames: obj(x.studentNames),
    }));
  // version 1: '전입·공동교육과정 이수내역' per student → kept as ⑥ entries so nothing is lost
  Object.entries(obj(raw.transferCredits)).forEach(([studentKey, list]) =>
    arr(list).forEach((t) => {
      if (!t || !t.name) return;
      extraCourses.push({
        id: t.id ? str(t.id) : newId("ex"),
        kind: "이전 버전 입력",
        name: str(t.name),
        group: str(t.group),
        credit: Number(t.credit) || 0,
        semester: str(t.semester),
        memo: "이전 버전의 ‘전입·공동교육과정 이수내역’에서 옮겨 온 항목입니다. 전입생의 전적교 과목이면 ⑤ 전입생 탭에 다시 입력하고 이 항목은 지워주세요.",
        studentKeys: [studentKey],
        studentNames: {
          [studentKey]: nameFromStudentKey(studentKey),
        },
      });
    }),
  );
  const manualStudents = arr(raw.manualStudents)
    .filter((m) => m && m.id && m.name)
    .map((m) => {
      var _a, _b, _c;
      return {
        id: str(m.id),
        name: str(m.name),
        grade: (_a = m.grade) !== null && _a !== void 0 ? _a : "",
        classNum: (_b = m.classNum) !== null && _b !== void 0 ? _b : "",
        number: (_c = m.number) !== null && _c !== void 0 ? _c : "",
      };
    });
  const archivedStudents = {};
  Object.entries(obj(raw.archivedStudents)).forEach(([k, v]) => {
    const o = obj(v);
    archivedStudents[k] = {
      reason: str(o.reason),
      archivedAt: o.archivedAt || null,
      name: str(o.name) || nameFromStudentKey(k),
    };
  });
  // version 3까지는 ② 화면에서 고른 한 가지 방법의 자료만 점검에 썼습니다. 4부터는 학기마다 자료를 고르므로,
  // 예전 파일을 열어도 결과가 그대로이도록 그때 쓰지 않던 쪽 파일은 ‘계산에 사용’을 꺼 둡니다(선생님이 다시 켤 수 있음).
  const surveyFilesIn = normalizeSurveyFiles(raw.surveyFiles);
  if ((Number(raw.version) || 1) <= 3) {
    if (raw.rosterMode === "survey") regFiles.forEach((f) => (f.disabled = true));
    else surveyFilesIn.forEach((f) => (f.disabled = true));
  }
  const num = (v, fallback) =>
    v === null || v === undefined || v === "" || !Number.isFinite(Number(v)) ? fallback : Number(v);
  return {
    version: Number(raw.version) || 1,
    savedAt: raw.savedAt || null,
    curriculumFile,
    reqOverride: Object.fromEntries(
      Object.entries(curriculumReq || obj(raw.reqOverride)).map(([k, v]) => [k, Number(v) || 0]),
    ),
    minCredit: num(raw.minCredit, 174),
    regFiles,
    semesterSourceChoice: obj(raw.semesterSourceChoice),
    mergeGroups: arr(raw.mergeGroups).filter((g) => g && g.id && Array.isArray(g.recordIds)),
    archivedStudents,
    classTeachers: obj(raw.classTeachers),
    builderData: normalizeBuilderData(raw.builderData),
    priorityGroups:
      raw.priorityGroups && typeof raw.priorityGroups === "object" && !Array.isArray(raw.priorityGroups)
        ? raw.priorityGroups
        : {
            "예술": true,
            "생활교양": true,
          },
    admissionYear: num(raw.admissionYear, new Date().getFullYear()),
    creativeActivityCredit: num(raw.creativeActivityCredit, 18),
    classMappingFiles: arr(raw.classMappingFiles)
      .filter((f) => f && typeof f === "object")
      .map((f) => ({
        ...f,
        entries: arr(f.entries),
      })),
    courseChanges,
    transferInfo,
    extraCourses,
    manualStudents,
    rosterMode: raw.rosterMode === "survey" ? "survey" : "upload",
    surveyConfig: normalizeSurveyConfig(raw.surveyConfig),
    surveyFiles: surveyFilesIn,
    grade1Electives: raw.grade1Electives === "none" || raw.grade1Electives === "some" ? raw.grade1Electives : null,
    surveyChoice: Object.fromEntries(Object.entries(obj(raw.surveyChoice)).map(([k, v]) => [k, str(v)])),
    surveyExcluded: Object.fromEntries(Object.entries(obj(raw.surveyExcluded)).filter(([, v]) => v === true)),
  };
}
// ---------- QR code (byte mode, offline) ----------
// 기초조사 주소를 반별 안내문에 인쇄하기 위한 QR 코드 생성기입니다. 인터넷 없이 이 파일 안에서 만듭니다.
const QR_ECC_CODEWORDS = {
  L: [
    -1, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30,
    30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30,
  ],
  M: [
    -1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28,
    28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28,
  ],
  Q: [
    -1, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30,
    30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30,
  ],
  H: [
    -1, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30,
    30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30,
  ],
};
const QR_ECC_BLOCKS = {
  L: [
    -1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19,
    19, 20, 21, 22, 24, 25,
  ],
  M: [
    -1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31,
    33, 35, 37, 38, 40, 43, 45, 47, 49,
  ],
  Q: [
    -1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43,
    45, 48, 51, 53, 56, 59, 62, 65, 68,
  ],
  H: [
    -1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48,
    51, 54, 57, 60, 63, 66, 70, 74, 77, 81,
  ],
};
const QR_ECC_FORMAT_BITS = {
  L: 1,
  M: 0,
  Q: 3,
  H: 2,
};
function qrRawModules(ver) {
  let result = (16 * ver + 128) * ver + 64;
  if (ver >= 2) {
    const numAlign = Math.floor(ver / 7) + 2;
    result -= (25 * numAlign - 10) * numAlign - 55;
    if (ver >= 7) result -= 36;
  }
  return result;
}
function qrDataCodewords(ver, ecl) {
  return Math.floor(qrRawModules(ver) / 8) - QR_ECC_CODEWORDS[ecl][ver] * QR_ECC_BLOCKS[ecl][ver];
}
function qrAlignPositions(ver) {
  if (ver === 1) return [];
  const numAlign = Math.floor(ver / 7) + 2;
  const step = ver === 32 ? 26 : Math.ceil((ver * 4 + 4) / (numAlign * 2 - 2)) * 2;
  const size = ver * 4 + 17;
  const result = [6];
  for (let pos = size - 7; result.length < numAlign; pos -= step) result.splice(1, 0, pos);
  return result;
}
function qrGfMul(x, y) {
  let z = 0;
  for (let i = 7; i >= 0; i--) {
    z = (z << 1) ^ ((z >>> 7) * 0x11d);
    z ^= ((y >>> i) & 1) * x;
  }
  return z & 0xff;
}
function qrRsDivisor(degree) {
  const result = new Array(degree).fill(0);
  result[degree - 1] = 1;
  let root = 1;
  for (let i = 0; i < degree; i++) {
    for (let j = 0; j < result.length; j++) {
      result[j] = qrGfMul(result[j], root);
      if (j + 1 < result.length) result[j] ^= result[j + 1];
    }
    root = qrGfMul(root, 0x02);
  }
  return result;
}
function qrRsRemainder(data, divisor) {
  const result = divisor.map(() => 0);
  data.forEach((b) => {
    const factor = b ^ result.shift();
    result.push(0);
    divisor.forEach((coef, i) => (result[i] ^= qrGfMul(coef, factor)));
  });
  return result;
}
function utf8Bytes(text) {
  const out = [];
  const s = String(text);
  for (let i = 0; i < s.length; i++) {
    let cp = s.codePointAt(i);
    if (cp > 0xffff) i++;
    if (cp < 0x80) out.push(cp);
    else if (cp < 0x800) out.push(0xc0 | (cp >> 6), 0x80 | (cp & 63));
    else if (cp < 0x10000) out.push(0xe0 | (cp >> 12), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
    else out.push(0xf0 | (cp >> 18), 0x80 | ((cp >> 12) & 63), 0x80 | ((cp >> 6) & 63), 0x80 | (cp & 63));
  }
  return out;
}
// returns { size, version, mask, modules: boolean[][] } — modules[y][x] true = dark
function qrEncodeText(text, opts) {
  const ecl = (opts && opts.ecl) || "M";
  const forcedVersion = opts && opts.version;
  const forcedMask = opts && opts.mask != null ? opts.mask : -1;
  const bytes = utf8Bytes(text);
  let ver = 0;
  for (let v = forcedVersion || 1; v <= (forcedVersion || 40); v++) {
    const countBits = v < 10 ? 8 : 16;
    if (4 + countBits + bytes.length * 8 <= qrDataCodewords(v, ecl) * 8) {
      ver = v;
      break;
    }
  }
  if (!ver) throw new Error("QR 코드에 담기에는 주소가 너무 깁니다.");
  const bits = [];
  const put = (val, len) => {
    for (let i = len - 1; i >= 0; i--) bits.push((val >>> i) & 1);
  };
  put(4, 4);
  put(bytes.length, ver < 10 ? 8 : 16);
  bytes.forEach((b) => put(b, 8));
  const capacity = qrDataCodewords(ver, ecl) * 8;
  put(0, Math.min(4, capacity - bits.length));
  put(0, (8 - (bits.length % 8)) % 8);
  for (let pad = 0xec; bits.length < capacity; pad ^= 0xec ^ 0x11) put(pad, 8);
  const data = [];
  for (let i = 0; i < bits.length; i += 8) {
    let b = 0;
    for (let j = 0; j < 8; j++) b = (b << 1) | bits[i + j];
    data.push(b);
  }
  // split into blocks, add error correction, interleave
  const numBlocks = QR_ECC_BLOCKS[ecl][ver];
  const eccLen = QR_ECC_CODEWORDS[ecl][ver];
  const rawCodewords = Math.floor(qrRawModules(ver) / 8);
  const numShort = numBlocks - (rawCodewords % numBlocks);
  const shortLen = Math.floor(rawCodewords / numBlocks);
  const divisor = qrRsDivisor(eccLen);
  const blocks = [];
  for (let i = 0, k = 0; i < numBlocks; i++) {
    const dat = data.slice(k, k + shortLen - eccLen + (i < numShort ? 0 : 1));
    k += dat.length;
    const ecc = qrRsRemainder(dat, divisor);
    if (i < numShort) dat.push(0); // placeholder so every block has the same length; skipped when interleaving
    blocks.push(dat.concat(ecc));
  }
  const codewords = [];
  for (let i = 0; i < blocks[0].length; i++) {
    blocks.forEach((blk, j) => {
      if (i !== shortLen - eccLen || j >= numShort) codewords.push(blk[i]);
    });
  }
  const size = ver * 4 + 17;
  const modules = Array.from(
    {
      length: size,
    },
    () => new Array(size).fill(false),
  );
  const isFn = Array.from(
    {
      length: size,
    },
    () => new Array(size).fill(false),
  );
  const setFn = (x, y, dark) => {
    modules[y][x] = dark;
    isFn[y][x] = true;
  };
  for (let i = 0; i < size; i++) {
    setFn(6, i, i % 2 === 0);
    setFn(i, 6, i % 2 === 0);
  }
  const finder = (cx, cy) => {
    for (let dy = -4; dy <= 4; dy++)
      for (let dx = -4; dx <= 4; dx++) {
        const dist = Math.max(Math.abs(dx), Math.abs(dy));
        const x = cx + dx;
        const y = cy + dy;
        if (x >= 0 && x < size && y >= 0 && y < size) setFn(x, y, dist !== 2 && dist !== 4);
      }
  };
  finder(3, 3);
  finder(size - 4, 3);
  finder(3, size - 4);
  const align = qrAlignPositions(ver);
  align.forEach((ay, i) =>
    align.forEach((ax, j) => {
      if ((i === 0 && j === 0) || (i === 0 && j === align.length - 1) || (i === align.length - 1 && j === 0)) return;
      for (let dy = -2; dy <= 2; dy++)
        for (let dx = -2; dx <= 2; dx++) setFn(ax + dx, ay + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
    }),
  );
  const drawFormat = (mask) => {
    const dataBits = (QR_ECC_FORMAT_BITS[ecl] << 3) | mask;
    let rem = dataBits;
    for (let i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
    const fbits = ((dataBits << 10) | rem) ^ 0x5412;
    const bit = (i) => ((fbits >>> i) & 1) !== 0;
    for (let i = 0; i <= 5; i++) setFn(8, i, bit(i));
    setFn(8, 7, bit(6));
    setFn(8, 8, bit(7));
    setFn(7, 8, bit(8));
    for (let i = 9; i < 15; i++) setFn(14 - i, 8, bit(i));
    for (let i = 0; i < 8; i++) setFn(size - 1 - i, 8, bit(i));
    for (let i = 8; i < 15; i++) setFn(8, size - 15 + i, bit(i));
    setFn(8, size - 8, true);
  };
  drawFormat(0); // reserve
  if (ver >= 7) {
    let rem = ver;
    for (let i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1f25);
    const vbits = (ver << 12) | rem;
    for (let i = 0; i < 18; i++) {
      const dark = ((vbits >>> i) & 1) !== 0;
      const a = size - 11 + (i % 3);
      const b = Math.floor(i / 3);
      setFn(a, b, dark);
      setFn(b, a, dark);
    }
  }
  // zigzag data placement
  let bitIndex = 0;
  const totalBits = codewords.length * 8;
  for (let right = size - 1; right >= 1; right -= 2) {
    if (right === 6) right = 5;
    for (let vert = 0; vert < size; vert++) {
      for (let j = 0; j < 2; j++) {
        const x = right - j;
        const upward = ((right + 1) & 2) === 0;
        const y = upward ? size - 1 - vert : vert;
        if (!isFn[y][x] && bitIndex < totalBits) {
          modules[y][x] = ((codewords[bitIndex >>> 3] >>> (7 - (bitIndex & 7))) & 1) !== 0;
          bitIndex++;
        }
      }
    }
  }
  const applyMask = (mask) => {
    for (let y = 0; y < size; y++)
      for (let x = 0; x < size; x++) {
        let invert;
        switch (mask) {
          case 0:
            invert = (x + y) % 2 === 0;
            break;
          case 1:
            invert = y % 2 === 0;
            break;
          case 2:
            invert = x % 3 === 0;
            break;
          case 3:
            invert = (x + y) % 3 === 0;
            break;
          case 4:
            invert = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0;
            break;
          case 5:
            invert = ((x * y) % 2) + ((x * y) % 3) === 0;
            break;
          case 6:
            invert = (((x * y) % 2) + ((x * y) % 3)) % 2 === 0;
            break;
          default:
            invert = (((x + y) % 2) + ((x * y) % 3)) % 2 === 0;
            break;
        }
        if (!isFn[y][x] && invert) modules[y][x] = !modules[y][x];
      }
  };
  const penalty = () => {
    let result = 0;
    const runPenalty = (line) => {
      let p = 0;
      let runColor = line[0];
      let runLen = 1;
      for (let i = 1; i <= line.length; i++) {
        if (i < line.length && line[i] === runColor) runLen++;
        else {
          if (runLen >= 5) p += runLen - 2;
          if (i < line.length) {
            runColor = line[i];
            runLen = 1;
          }
        }
      }
      // finder-like 1:1:3:1:1 with 4 light modules on either side (light outside the symbol counts)
      const at = (i) => (i >= 0 && i < line.length ? line[i] : false);
      for (let i = 0; i + 6 < line.length; i++) {
        if (at(i) && !at(i + 1) && at(i + 2) && at(i + 3) && at(i + 4) && !at(i + 5) && at(i + 6)) {
          const before = !at(i - 1) && !at(i - 2) && !at(i - 3) && !at(i - 4);
          const after = !at(i + 7) && !at(i + 8) && !at(i + 9) && !at(i + 10);
          if (before || after) p += 40;
        }
      }
      return p;
    };
    for (let y = 0; y < size; y++) result += runPenalty(modules[y]);
    for (let x = 0; x < size; x++) result += runPenalty(modules.map((row) => row[x]));
    for (let y = 0; y < size - 1; y++)
      for (let x = 0; x < size - 1; x++) {
        const c = modules[y][x];
        if (c === modules[y][x + 1] && c === modules[y + 1][x] && c === modules[y + 1][x + 1]) result += 3;
      }
    let dark = 0;
    modules.forEach((row) => row.forEach((c) => c && dark++));
    const total = size * size;
    const k = Math.ceil(Math.abs(dark * 20 - total * 10) / total) - 1;
    result += k * 10;
    return result;
  };
  let mask = forcedMask;
  if (mask < 0) {
    let best = Infinity;
    for (let m = 0; m < 8; m++) {
      applyMask(m);
      drawFormat(m);
      const p = penalty();
      if (p < best) {
        best = p;
        mask = m;
      }
      applyMask(m); // undo (xor)
    }
  }
  applyMask(mask);
  drawFormat(mask);
  return {
    size,
    version: ver,
    mask,
    modules,
  };
}
// one SVG path for all dark modules, with a 4-module quiet zone
function qrSvgPath(qr) {
  const parts = [];
  for (let y = 0; y < qr.size; y++) {
    let x = 0;
    while (x < qr.size) {
      if (!qr.modules[y][x]) {
        x++;
        continue;
      }
      const start = x;
      while (x < qr.size && qr.modules[y][x]) x++;
      parts.push(`M${start + 4} ${y + 4}h${x - start}v1h${start - x}z`);
    }
  }
  return parts.join("");
}
// ---------- 학생 기초조사: 설정 · 응답 파일 읽기 (pure) ----------
// 수강신청 파일 형식이 학교마다 달라 인식이 어긋날 때를 위한 두 번째 방법입니다. 학생이 휴대폰으로 자기 반·번호·이름과
// 학기별로 들은(신청한) 과목을 고르면 선생님의 구글 시트에 쌓이고, 그 시트를 엑셀로 내려받아 여기에 올립니다.
const SURVEY_AWAY_MARK = "(다른 학교에 다님)";
const SURVEY_META_LABEL = "조사 정보(수정하지 마세요)";
const SURVEY_CONTROL_RE = new RegExp(
  "[" +
    String.fromCharCode(0) +
    "-" +
    String.fromCharCode(9) +
    String.fromCharCode(11) +
    "-" +
    String.fromCharCode(31) +
    String.fromCharCode(127, 0x2028, 0x2029) +
    "]",
  "g",
);
const SURVEY_URL_RE = new RegExp(
  "^https://script[.]google[.]com/(a/macros/[^/?# ]+|macros)/s/([A-Za-z0-9_-]{20,})/(exec|dev)/?([?][^# ]*)?$",
);
function pad2(n) {
  return String(n).padStart(2, "0");
}
function fnv1aHex(s) {
  let h = 0x811c9dc5;
  for (let i = 0; i < s.length; i++) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}
// 조사 화면·시트에 들어갈 글자: 제어 문자 제거, 시트가 수식으로 읽는 첫 글자(=,+,-,@) 제거
function surveyText(v, max) {
  return String(v == null ? "" : v)
    .replace(SURVEY_CONTROL_RE, " ")
    .replace(/^[=+@-]+/, "")
    .trim()
    .slice(0, max);
}
function normalizePersonName(v) {
  return String(v == null ? "" : v)
    .replace(SURVEY_CONTROL_RE, " ")
    .replace(/^'/, "")
    .replace(/\s+/g, " ")
    .trim();
}
function guessCurrentTerm(now) {
  const m = (now || new Date()).getMonth() + 1;
  return m >= 3 && m <= 7 ? 1 : 2;
}
// 입학 연도로 지금 학년을 추정 (1~2월은 아직 앞 학년도)
function defaultSurveyGrade(admissionYear, now) {
  const d = now || new Date();
  const schoolYear = d.getMonth() + 1 >= 3 ? d.getFullYear() : d.getFullYear() - 1;
  const g = schoolYear - Number(admissionYear) + 1;
  return g >= 1 && g <= 3 ? g : 2;
}
function formatDateYmd(d) {
  const x = d || new Date();
  return `${x.getFullYear()}-${pad2(x.getMonth() + 1)}-${pad2(x.getDate())}`;
}
// 결번: [7, 15] 또는 "7, 15번" 같은 글자 → 1~60 사이의 번호 목록(중복 없이, 작은 순)
function parseSkipNumbers(v) {
  const list = Array.isArray(v) ? v : String(v == null ? "" : v).split(/[^0-9]+/);
  return Array.from(new Set(list.map((x) => Math.floor(Number(x))).filter((n) => n >= 1 && n <= 60))).sort(
    (a, b) => a - b,
  );
}
// 시트가 응답을 내보낼 때 확인하는 열쇠. 설정 코드에 실려 시트로 가고, 학생 화면에는 나가지 않습니다.
function newExportKey() {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
  const bytes = new Uint8Array(24);
  if (typeof crypto !== "undefined" && crypto.getRandomValues) crypto.getRandomValues(bytes);
  else for (let i = 0; i < bytes.length; i++) bytes[i] = Math.floor(Math.random() * 256);
  return Array.from(bytes, (b) => alphabet[b % alphabet.length]).join("");
}
function emptySurveyConfig() {
  return {
    id: newId("sv"),
    title: "",
    grade: null,
    term: null,
    semesters: null,
    classes: [],
    notice: "",
    deadline: "",
    webAppUrl: "",
    copiedFp: "",
    installMethod: "",
    installChecks: {},
    exportKey: newExportKey(),
  };
}
function normalizeSurveyConfig(raw) {
  const base = emptySurveyConfig();
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) return base;
  const str = (v, max) => String(v == null ? "" : v).slice(0, max);
  const seen = new Set();
  const classes = (Array.isArray(raw.classes) ? raw.classes : [])
    .map((k) => ({
      c: Math.floor(Number(k && k.c)),
      size: Math.floor(Number(k && k.size)) || 0,
      skip: k && k.skip,
    }))
    .filter((k) => k.c >= 1 && k.c <= 30 && !seen.has(k.c) && seen.add(k.c))
    .map((k) => ({
      c: k.c,
      size: Math.max(0, Math.min(60, k.size)),
      skip: parseSkipNumbers(k.skip),
    }))
    .sort((a, b) => a.c - b.c);
  return {
    id: typeof raw.id === "string" && /^[A-Za-z0-9_-]{1,40}$/.test(raw.id) ? raw.id : base.id,
    title: str(raw.title, 60),
    grade: [1, 2, 3].includes(Number(raw.grade)) ? Number(raw.grade) : null,
    term: [1, 2].includes(Number(raw.term)) ? Number(raw.term) : null,
    semesters: Array.isArray(raw.semesters) ? raw.semesters.filter((s) => SEMESTER_CODES.includes(s)) : null,
    classes,
    notice: str(raw.notice, 400),
    deadline: str(raw.deadline, 60),
    webAppUrl: str(raw.webAppUrl, 400),
    copiedFp: str(raw.copiedFp, 40),
    exportKey:
      typeof raw.exportKey === "string" && /^[A-Za-z0-9_-]{16,64}$/.test(raw.exportKey)
        ? raw.exportKey
        : base.exportKey,
    installMethod: ["template", "code"].includes(raw.installMethod) ? raw.installMethod : "",
    installChecks:
      raw.installChecks && typeof raw.installChecks === "object"
        ? Object.fromEntries(
            Object.keys(raw.installChecks)
              .filter((k) => /^[a-z][0-9]$/.test(k) && raw.installChecks[k] === true)
              .map((k) => [k, true]),
          )
        : {},
  };
}
// 편제표의 학생선택 묶음(택N)으로 학생 화면에 보여줄 조사 내용을 만듭니다.
function surveySemestersWithGroups(poolDefs) {
  return Array.from(new Set(poolDefs.map((d) => d.sem)))
    .filter((s) => SEMESTER_CODES.includes(s))
    .sort(compareSemester);
}
function buildSurveyPayload(config, poolDefs, opts) {
  const now = (opts && opts.now) || new Date();
  const grade = config.grade || defaultSurveyGrade(opts && opts.admissionYear, now);
  const term = config.term || guessCurrentTerm(now);
  const nowIdx = (grade - 1) * 2 + (term - 1);
  const available = surveySemestersWithGroups(poolDefs);
  const chosen = config.semesters ? available.filter((s) => config.semesters.includes(s)) : available;
  const semesters = chosen.map((code) => {
    const idx = SEMESTER_CODES.indexOf(code);
    const groups = poolDefs
      .filter((d) => d.sem === code)
      .map((d, i) => {
        const names = new Set();
        const options = [];
        d.courses.forEach((c) => {
          const name = surveyText(c.name, 60);
          if (!name || names.has(name)) return;
          names.add(name);
          options.push({
            name,
            group: toCurriculumGroup(c.group) || "",
            credit: Number(c.credit) || 0,
          });
        });
        return {
          id: `${code}#${i}`,
          label: surveyText(d.label, 40) || `선택 묶음 ${i + 1}`,
          n: d.requiredN != null ? d.requiredN : null,
          options,
        };
      });
    return {
      code,
      label: formatSemester(code),
      status: idx < nowIdx ? "done" : idx === nowIdx ? "now" : "next",
      groups,
    };
  });
  // skip은 결번이 있을 때만 넣습니다 — 결번이 없는 설정은 이전 버전과 코드 버전(fp)이 같게 유지됩니다
  const classes = config.classes
    .filter((k) => k.size >= 1)
    .map((k) => {
      const skip = (k.skip || []).filter((n) => n <= k.size);
      return skip.length
        ? {
            c: k.c,
            size: k.size,
            skip,
          }
        : {
            c: k.c,
            size: k.size,
          };
    });
  const content = {
    id: config.id,
    title: surveyText(config.title, 60) || "선택과목 기초조사",
    grade,
    term,
    notice: surveyText(config.notice, 400),
    deadline: surveyText(config.deadline, 60),
    classes,
    semesters,
  };
  return {
    v: 1,
    ...content,
    fp: "v-" + fnv1aHex(JSON.stringify(content)),
    createdAt: formatDateYmd(now),
    key: config.exportKey || undefined,
  };
}
function surveyProblems(payload, config, hasCurriculum) {
  const errors = [];
  const warnings = [];
  if (!hasCurriculum)
    errors.push("편제표를 먼저 올려주세요. 학생 화면의 과목 목록은 편제표의 학생선택 과목으로 만들어집니다.");
  if (!config.grade || !config.term)
    errors.push(
      "1번 ‘조사하는 때’에서 지금 학년과 학기를 골라주세요. 반·번호를 몇 학년 기준으로 받을지, 학기마다 ‘들은 과목/신청한 과목’ 중 무엇으로 보일지가 정해집니다.",
    );
  else if (payload.semesters.length === 0)
    errors.push("조사할 학기를 하나 이상 골라주세요. (편제표에 학생선택 과목이 있는 학기만 고를 수 있습니다)");
  payload.semesters.forEach((s) =>
    s.groups.forEach((g) => {
      if (g.n != null && g.n > g.options.length)
        errors.push(
          `${s.label} ${g.label}: 택${g.n}인데 고를 수 있는 과목이 ${g.options.length}개뿐입니다. 편제표를 확인해주세요.`,
        );
      if (g.n == null) warnings.push(`${s.label} ‘${g.label}’에는 택N 표시가 없어서 학생이 원하는 만큼 고르게 됩니다.`);
    }),
  );
  if (config.classes.length === 0) errors.push("반 수와 반별 인원을 입력해주세요. (누가 안 냈는지 확인하는 데 씁니다)");
  config.classes.forEach((k) => {
    if (!(k.size >= 1 && k.size <= 60)) errors.push(`${k.c}반 인원을 1~60 사이 숫자로 입력해주세요.`);
    else if ((k.skip || []).filter((n) => n <= k.size).length >= k.size)
      errors.push(`${k.c}반은 모든 번호가 결번으로 되어 있습니다. 결번을 확인해주세요.`);
  });
  return {
    errors,
    warnings,
  };
}
function checkWebAppUrl(raw) {
  const url = String(raw || "").trim();
  if (!url)
    return {
      ok: false,
      empty: true,
      message: "",
    };
  const m = SURVEY_URL_RE.exec(url);
  if (!m)
    return {
      ok: false,
      message:
        "구글 Apps Script 웹 앱 주소가 아닙니다. [배포 → 새 배포]를 마친 뒤 나오는 ‘웹 앱 URL’(https://script.google.com/macros/s/…/exec)을 복사해 붙여넣어주세요.",
    };
  if (m[3] === "dev")
    return {
      ok: false,
      message:
        "…/dev 로 끝나는 주소는 선생님만 열 수 있는 테스트 주소입니다. [배포 → 새 배포]에서 받은 …/exec 주소를 넣어주세요.",
    };
  const q = url.indexOf("?");
  return {
    ok: true,
    url: (q >= 0 ? url.slice(0, q) : url).replace(/[/]$/, ""),
  };
}
function surveyClassUrl(baseUrl, c) {
  return `${baseUrl}?ban=${c}`;
}
// ---- 응답 파일 (구글 시트 → 파일 → 다운로드 → Microsoft Excel) ----
function findHeaderRow(rows, required) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const r = (rows[i] || []).map((c) => String(c == null ? "" : c).trim());
    if (required.every((h) => r.includes(h))) {
      const cols = {};
      r.forEach((h, j) => {
        if (h && cols[h] == null) cols[h] = j;
      });
      return {
        index: i,
        cols,
      };
    }
  }
  return null;
}
function parseSurveyWorkbook(wb) {
  let picks = null;
  let log = null;
  let meta = null;
  wb.SheetNames.forEach((sn) => {
    const rows = sheetToMatrix(wb, sn) || [];
    if (!picks) {
      const h = findHeaderRow(rows, ["제출번호", "반", "번호", "이름", "학기", "선택그룹", "과목"]);
      if (h)
        picks = {
          rows,
          h,
        };
    }
    if (!log) {
      const h = findHeaderRow(rows, ["제출번호", "제출시각", "남긴 말"]);
      if (h)
        log = {
          rows,
          h,
        };
    }
    if (!meta)
      rows.forEach((r) => {
        if (meta || !r || String(r[0] == null ? "" : r[0]).trim() !== SURVEY_META_LABEL || !r[1]) return;
        try {
          const m = JSON.parse(String(r[1]));
          if (m && m.kind === "graduation-survey") meta = m;
        } catch (e) {
          // 설정 시트를 손으로 고친 경우 — 조사 정보 없이 계속
        }
      });
  });
  if (!picks) {
    if (detectRegistrationFormat(wb))
      throw new Error("수강신청 결과 파일입니다. 이 파일은 위의 ‘수강신청 엑셀 업로드’를 골라서 올려주세요.");
    throw new Error(
      "기초조사 응답 파일이 아닙니다. 조사용 구글 시트에서 [파일 → 다운로드 → Microsoft Excel(.xlsx)]로 내려받은 파일을 올려주세요. (‘선택과목’ 시트의 제출번호·학기·선택그룹·과목 머리글을 찾지 못했습니다)",
    );
  }
  const text = (v) =>
    String(v == null ? "" : v)
      .replace(/^'/, "")
      .trim();
  const cell = (row, name) => {
    const j = picks.h.cols[name];
    return j == null ? null : row[j];
  };
  const bySub = new Map();
  let skipped = 0;
  for (let i = picks.h.index + 1; i < picks.rows.length; i++) {
    const row = picks.rows[i];
    if (!row || !row.some((c) => c != null && String(c).trim() !== "")) continue;
    const no = Number(cell(row, "제출번호"));
    const grade = Number(cell(row, "학년"));
    const classNum = Number(cell(row, "반"));
    const number = Number(cell(row, "번호"));
    const name = normalizePersonName(cell(row, "이름"));
    const sem = normalizeSemesterCode(cell(row, "학기"));
    const subject = text(cell(row, "과목"));
    if (!(no > 0) || !(classNum > 0) || !(number > 0) || !name || !SEMESTER_CODES.includes(sem) || !subject) {
      skipped++;
      continue;
    }
    const k = String(no);
    if (!bySub.has(k))
      bySub.set(k, {
        no,
        at: text(cell(row, "제출시각")),
        grade: grade >= 1 && grade <= 3 ? grade : null,
        classNum,
        number,
        name,
        subjects: [],
        away: [],
        memo: "",
      });
    const s = bySub.get(k);
    if (subject === SURVEY_AWAY_MARK) {
      if (!s.away.includes(sem)) s.away.push(sem);
    } else if (!s.subjects.some((x) => x.name === subject && x.semester === sem)) {
      s.subjects.push({
        name: subject,
        semester: sem,
      });
    }
  }
  if (log) {
    for (let i = log.h.index + 1; i < log.rows.length; i++) {
      const row = log.rows[i];
      if (!row) continue;
      const s = bySub.get(String(Number(row[log.h.cols["제출번호"]])));
      if (s) s.memo = text(row[log.h.cols["남긴 말"]]).slice(0, 300);
    }
  }
  return {
    meta,
    submissions: Array.from(bySub.values()).sort((a, b) => a.no - b.no),
    skipped,
  };
}
function normalizeSurveyFiles(list) {
  const arr = (v) => (Array.isArray(v) ? v : []);
  return arr(list)
    .filter((f) => f && typeof f === "object")
    .map((f, i) => ({
      id: f.id ? String(f.id) : `sf-${i}`,
      name: String(f.name == null ? "" : f.name),
      status: f.status === "error" ? "error" : "ok",
      error: f.error ? String(f.error) : "",
      meta: f.meta && typeof f.meta === "object" && !Array.isArray(f.meta) ? f.meta : null,
      skipped: Number(f.skipped) || 0,
      disabled: f.disabled === true,
      addedAt: Number(f.addedAt) || 0,
      submissions: arr(f.submissions)
        .filter((s) => s && typeof s === "object")
        .map((s) => ({
          no: Number(s.no) || 0,
          at: String(s.at == null ? "" : s.at),
          grade: [1, 2, 3].includes(Number(s.grade)) ? Number(s.grade) : null,
          classNum: Number(s.classNum) || 0,
          number: Number(s.number) || 0,
          name: normalizePersonName(s.name),
          subjects: arr(s.subjects)
            .filter((x) => x && x.name && SEMESTER_CODES.includes(x.semester))
            .map((x) => ({
              name: String(x.name),
              semester: x.semester,
            })),
          away: arr(s.away).filter((c) => SEMESTER_CODES.includes(c)),
          memo: String(s.memo == null ? "" : s.memo),
        }))
        .filter((s) => s.no > 0 && s.classNum > 0 && s.number > 0 && s.name),
    }));
}
function surveyIdentity(s) {
  return `${s.grade || 0}-${s.classNum}-${s.number}`;
}
function surveyIdText(s) {
  return `${s.grade ? s.grade + "학년 " : ""}${s.classNum}반 ${s.number}번`;
}
// 올린 응답 파일들 → 학생 한 명당 기록 하나 (같은 번호로 여러 번 냈으면 마지막 제출, 또는 선생님이 고른 제출)
function buildSurveyData(surveyFiles, opts) {
  const choice = (opts && opts.choice) || {};
  const excluded = (opts && opts.excluded) || {};
  const classes = (opts && opts.classes) || [];
  const grade = (opts && opts.grade) || null;
  const all = [];
  const usedFiles = surveyFiles.filter((f) => f.status === "ok" && !f.disabled); // ‘계산에 사용’을 끈 파일은 뺍니다
  usedFiles.forEach((f) => {
    (f.submissions || []).forEach((s) =>
      all.push({
        ...s,
        fileId: f.id,
        fileName: f.name,
        key: `${f.id}#${s.no}`,
      }),
    );
  });
  // 같은 시트를 여러 번 내려받아 모두 올린 경우 같은 제출이 겹치므로 하나만 남깁니다 (나중에 올린 파일 쪽)
  const bySig = new Map();
  all.forEach((s) => bySig.set(`${s.no}|${s.at}|${surveyIdentity(s)}|${s.name}`, s));
  const unique = Array.from(bySig.values());
  const usable = unique.filter((s) => !excluded[s.key]);
  const order = (a, b) => String(a.at).localeCompare(String(b.at)) || a.no - b.no;
  const byId = new Map();
  usable.forEach((s) => {
    const id = surveyIdentity(s);
    if (!byId.has(id)) byId.set(id, []);
    byId.get(id).push(s);
  });
  const records = [];
  const conflicts = [];
  const resubmitted = [];
  Array.from(byId.entries())
    .sort((a, b) => a[1][0].classNum - b[1][0].classNum || a[1][0].number - b[1][0].number)
    .forEach(([id, subs]) => {
      subs.sort(order);
      const picked = subs.find((s) => s.key === choice[id]) || subs[subs.length - 1];
      const names = Array.from(new Set(subs.map((s) => s.name)));
      if (names.length > 1)
        conflicts.push({
          id,
          idText: surveyIdText(picked),
          subs,
          pickedKey: picked.key,
        });
      else if (subs.length > 1)
        resubmitted.push({
          id,
          idText: surveyIdText(picked),
          name: picked.name,
          count: subs.length,
        });
      records.push({
        uid: "sv:" + picked.key,
        fileId: picked.fileId,
        source: "survey",
        fileLabel: "기초조사",
        name: picked.name,
        idFields: {
          grade: picked.grade,
          classNum: picked.classNum,
          number: picked.number,
        },
        idKey: `${picked.grade || 0}${pad2(picked.classNum)}${pad2(picked.number)}`,
        idText: surveyIdText(picked),
        subjects: picked.subjects.map((x) => ({
          name: x.name,
          semester: x.semester,
        })),
        semesters: Array.from(new Set(picked.subjects.map((x) => x.semester))).sort(compareSemester),
        blockId: "S",
        away: picked.away,
        memo: picked.memo,
        submittedAt: picked.at,
        submitNo: picked.no,
        submitKey: picked.key,
        submitCount: subs.length,
      });
    });
  const byName = new Map();
  records.forEach((r) => {
    if (!byName.has(r.name)) byName.set(r.name, []);
    byName.get(r.name).push(r);
  });
  const sameName = Array.from(byName.entries())
    .filter(([, rs]) => rs.length > 1)
    .map(([name, rs]) => ({
      name,
      records: rs,
    }));
  const submittedIds = new Set(
    records
      .filter((r) => !grade || r.idFields.grade === grade)
      .map((r) => `${r.idFields.classNum}-${r.idFields.number}`),
  );
  const missingByClass = classes.map((k) => {
    const skip = (k.skip || []).filter((n) => n <= k.size);
    const missing = [];
    for (let n = 1; n <= k.size; n++) if (!skip.includes(n) && !submittedIds.has(`${k.c}-${n}`)) missing.push(n);
    const size = k.size - skip.length; // 결번을 뺀 실제 인원
    return {
      c: k.c,
      size,
      submitted: size - missing.length,
      missing,
      skip,
    };
  });
  const outside =
    classes.length === 0
      ? []
      : records.filter(
          (r) =>
            (grade && r.idFields.grade !== grade) ||
            !classes.some(
              (k) =>
                k.c === r.idFields.classNum &&
                r.idFields.number <= k.size &&
                !(k.skip || []).includes(r.idFields.number),
            ),
        );
  const coveredSemesters = new Set();
  usedFiles.forEach((f) => {
    if (f.meta && Array.isArray(f.meta.semesters))
      f.meta.semesters.forEach((s) => SEMESTER_CODES.includes(s) && coveredSemesters.add(s));
  });
  // 학기마다 기초조사에 답한 학생 수 (그 학기 과목을 골랐거나 ‘다른 학교’로 표시)
  const semesterCounts = new Map();
  records.forEach((r) =>
    new Set([...r.semesters, ...r.away]).forEach((s) => semesterCounts.set(s, (semesterCounts.get(s) || 0) + 1)),
  );
  records.forEach((r) => {
    r.semesters.forEach((s) => coveredSemesters.add(s));
    r.away.forEach((s) => coveredSemesters.add(s));
  });
  return {
    records,
    conflicts,
    resubmitted,
    sameName,
    ambiguousNames: new Set(sameName.map((x) => x.name)),
    missingByClass,
    missingTotal: missingByClass.reduce((a, k) => a + k.missing.length, 0),
    outside,
    transferMarked: records.filter((r) => r.away.length > 0),
    memos: records.filter((r) => r.memo),
    coveredSemesters,
    totalSubmissions: unique.length,
    excludedSubs: unique.filter((s) => excluded[s.key]),
    duplicateAcrossFiles: all.length - unique.length,
    semesterCounts,
    lastAddedAt: usedFiles.reduce((a, f) => Math.max(a, Number(f.addedAt) || 0), 0),
  };
}
// 학기별 자료 출처에서 ‘학생 기초조사’를 고른 학기의 과목만 남깁니다. 한 학기에는 자료 하나만 쓰므로,
// 다른 학기는 수강신청 파일 쪽 기록에서 채워지고 같은 학생끼리는 이름(동명이인이면 학년·반·번호)으로 합쳐집니다.
const SURVEY_SOURCE_ID = "survey";
function buildSurveyRawRecords(surveyData, chosenSourceBySem) {
  const sems = new Set(
    Array.from(chosenSourceBySem.entries())
      .filter(([, id]) => id === SURVEY_SOURCE_ID)
      .map(([sem]) => sem),
  );
  if (sems.size === 0) return [];
  return surveyData.records.map((r) => {
    const subjects = r.subjects.filter((s) => sems.has(s.semester));
    return {
      ...r,
      subjects,
      semesters: Array.from(new Set(subjects.map((s) => s.semester))).sort(compareSemester),
      away: r.away.filter((s) => sems.has(s)),
    };
  });
}
// 1학년 1·2학기의 학생선택 과목: 선생님이 고른 답과 편제표를 맞춰 봅니다.
// answer: 'none'(모두 학교지정) | 'some'(학생선택 있음) | null(아직 안 고름)
function grade1ElectiveStatus(poolDefs, answer) {
  const pools = poolDefs.filter((d) => d.sem === "1-1" || d.sem === "1-2");
  const has = pools.length > 0;
  const picked = answer === "none" || answer === "some" ? answer : null;
  return {
    has,
    sems: Array.from(new Set(pools.map((d) => d.sem))).sort(compareSemester),
    groupCount: pools.length,
    courseCount: pools.reduce((a, d) => a + d.courses.length, 0),
    answer: picked,
    level: !picked ? "ask" : (picked === "some") === has ? "ok" : "warn",
  };
}
// ---------- ④ 변경의 수강신청 시스템 반영 · ③ 확인 사유 분류 (pure) ----------
// 반영 표시는 그때의 변경 내역 ‘서명’과 함께 저장합니다. 표시한 뒤 과목을 또 바꾸면 서명이 달라져 ‘반영 후 다시 바뀜’이 됩니다.
function changeSignature(change) {
  if (!change) return "";
  const removed = (Array.isArray(change.removed) ? change.removed : []).map(String).sort();
  const added = (Array.isArray(change.added) ? change.added : [])
    .filter((a) => a && a.name)
    .map((a) => entryKey(a.name, a.semester))
    .sort();
  return `-${removed.join(",")}|+${added.join(",")}`;
}
const REASON_CATEGORY_LABELS = {
  total: "총 학점 미달",
  dup: "같은 과목 두 학기 선택",
  missing: "신청 기록 없는 학기",
  pool: "택N 개수 안 맞음",
  mismatch: "편제표와 학기 다른 과목",
  unmatched: "인식 못한 과목",
};
function reasonCategory(reason) {
  if (reason.type === "short") {
    const group = reason.group || String(reason.text || "").split(" ")[0];
    return {
      key: "short:" + group,
      label: `${group} 학점 부족`,
      blocking: true,
    };
  }
  return {
    key: String(reason.type),
    label: REASON_CATEGORY_LABELS[reason.type] || String(reason.text || ""),
    blocking: !!reason.blocking,
  };
}
// ---------- small UI pieces ----------
function UploadCard({ title, desc, multiple, onFiles, files, onRemove, onRelabel }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const fs = Array.from(e.dataTransfer.files || []);
        if (fs.length) onFiles(fs);
      }}
      style={{
        border: `1.5px dashed ${dragOver ? ACCENT : LINE}`,
        borderRadius: 10,
        padding: 20,
        background: dragOver ? ACCENT_BG : "#FFFFFF",
        transition: "background 0.15s, border-color 0.15s",
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "flex-start",
          gap: 12,
        }}
      >
        <div
          style={{
            width: 36,
            height: 36,
            borderRadius: 8,
            background: PAPER,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          <FileSpreadsheet size={18} color={ACCENT} />
        </div>
        <div
          style={{
            flex: 1,
            minWidth: 0,
          }}
        >
          <div
            style={{
              fontWeight: 600,
              fontSize: 15,
              color: INK,
            }}
          >
            {title}
          </div>
          <div
            style={{
              fontSize: 13,
              color: MUTED,
              marginTop: 2,
              lineHeight: 1.5,
            }}
          >
            {desc}
          </div>
          <div
            style={{
              marginTop: 12,
              display: "flex",
              gap: 8,
              flexWrap: "wrap",
            }}
          >
            <button
              onClick={() => {
                var _a;
                return (_a = inputRef.current) === null || _a === void 0 ? void 0 : _a.click();
              }}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                fontSize: 13,
                fontWeight: 600,
                color: "#fff",
                background: ACCENT,
                border: "none",
                borderRadius: 7,
                padding: "8px 14px",
                cursor: "pointer",
              }}
            >
              <Upload size={14} />
              파일 선택
            </button>
            <span
              style={{
                fontSize: 12,
                color: MUTED,
                alignSelf: "center",
              }}
            >
              또는 여기에 끌어다 놓기
            </span>
          </div>
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx,.xls"
            multiple={multiple}
            onChange={(e) => {
              const fs = Array.from(e.target.files || []);
              if (fs.length) onFiles(fs);
              e.target.value = "";
            }}
            style={{
              display: "none",
            }}
          />
        </div>
      </div>
      {files && files.length > 0 && (
        <div
          style={{
            marginTop: 14,
            display: "flex",
            flexDirection: "column",
            gap: 6,
          }}
        >
          {files.map((f, idx) => {
            var _a;
            return (
              <div
                key={idx}
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  fontSize: 12.5,
                  background: PAPER,
                  borderRadius: 6,
                  padding: "6px 10px",
                  gap: 8,
                }}
              >
                <div
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    minWidth: 0,
                    flex: 1,
                  }}
                >
                  {f.status === "error" ? (
                    <AlertCircle
                      size={14}
                      color={WARN}
                      style={{
                        flexShrink: 0,
                      }}
                    />
                  ) : (
                    <CheckCircle2
                      size={14}
                      color={OK}
                      style={{
                        flexShrink: 0,
                      }}
                    />
                  )}
                  {onRelabel ? (
                    <input
                      value={(_a = f.label) !== null && _a !== void 0 ? _a : f.name}
                      onChange={(e) => onRelabel(idx, e.target.value)}
                      style={{
                        border: "none",
                        background: "transparent",
                        fontSize: 12.5,
                        color: INK,
                        flex: 1,
                        minWidth: 0,
                        padding: "2px 0",
                      }}
                    />
                  ) : (
                    <span
                      style={{
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                        color: INK,
                      }}
                    >
                      {f.name}
                    </span>
                  )}
                  {f.count != null && (
                    <span
                      style={{
                        color: MUTED,
                        flexShrink: 0,
                      }}
                    >
                      {"\u00B7 "}
                      {f.count}명
                    </span>
                  )}
                  {f.status === "error" && (
                    <span
                      style={{
                        color: WARN,
                        flexShrink: 0,
                      }}
                    >
                      {"\u00B7 "}
                      {f.error}
                    </span>
                  )}
                </div>
                <button
                  onClick={() => onRemove(idx)}
                  style={{
                    border: "none",
                    background: "transparent",
                    cursor: "pointer",
                    color: MUTED,
                    flexShrink: 0,
                  }}
                >
                  <X size={14} />
                </button>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}
function StatCard({ label, value, tone, onClick, active, hint }) {
  const body = (
    <>
      <div
        style={{
          fontSize: 28,
          fontWeight: 700,
          color: tone || INK,
          fontVariantNumeric: "tabular-nums",
        }}
      >
        {value}
      </div>
      <div
        style={{
          fontSize: 12.5,
          color: MUTED,
          marginTop: 4,
        }}
      >
        {label}
      </div>
      {onClick && (
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            marginTop: 4,
            color: active ? ACCENT : "#A8A395",
          }}
        >
          {active ? "▼ 아래에 표시 중 (다시 누르면 닫힘)" : hint || "눌러서 보기"}
        </div>
      )}
    </>
  );
  const style = {
    background: active ? ACCENT_BG : "#fff",
    border: `${active ? 2 : 1}px solid ${active ? ACCENT : LINE}`,
    borderRadius: 10,
    padding: active ? "15px 19px" : "16px 20px",
    flex: 1,
    minWidth: 140,
  };
  if (!onClick) return <div style={style}>{body}</div>;
  return (
    <button
      onClick={onClick}
      style={{
        ...style,
        textAlign: "left",
        cursor: "pointer",
        fontFamily: FONT,
        color: INK,
      }}
    >
      {body}
    </button>
  );
}
function Th({ children }) {
  return (
    <th
      style={{
        textAlign: "left",
        padding: "10px 12px",
        fontSize: 12,
        fontWeight: 700,
        color: MUTED,
        whiteSpace: "nowrap",
        position: "sticky",
        top: 0,
        background: PAPER,
      }}
    >
      {children}
    </th>
  );
}
function Td({ children, style }) {
  return (
    <td
      style={{
        padding: "9px 12px",
        whiteSpace: "nowrap",
        ...style,
      }}
    >
      {children}
    </td>
  );
}
// Catches render crashes anywhere below it and shows a readable message instead of a blank/frozen
// screen — if something like a malformed row in an uploaded file trips up a specific student's
// detail view, the rest of the app (and the error text needed to diagnose it) stays visible.
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      error: null,
    };
  }
  static getDerivedStateFromError(error) {
    return {
      error,
    };
  }
  render() {
    var _a;
    if (this.state.error) {
      return (
        <div
          style={{
            fontFamily: FONT,
            padding: 32,
            maxWidth: 700,
            margin: "0 auto",
          }}
        >
          <div
            style={{
              fontSize: 16,
              fontWeight: 700,
              color: WARN,
              marginBottom: 10,
            }}
          >
            문제가 발생했습니다
          </div>
          <div
            style={{
              fontSize: 13,
              color: MUTED,
              marginBottom: 14,
              lineHeight: 1.6,
            }}
          >
            화면을 표시하는 중 오류가 났어요. 아래 내용을 캡처해서 알려주시면 원인을 정확히 찾아 고칠 수 있습니다.
          </div>
          <pre
            style={{
              whiteSpace: "pre-wrap",
              fontSize: 12,
              background: PAPER,
              borderRadius: 8,
              padding: 14,
              color: INK,
              overflow: "auto",
            }}
          >
            {String(((_a = this.state.error) === null || _a === void 0 ? void 0 : _a.stack) || this.state.error)}
          </pre>
          <button
            onClick={() =>
              this.setState({
                error: null,
              })
            }
            style={{
              marginTop: 14,
              fontSize: 13,
              fontWeight: 700,
              color: "#fff",
              background: ACCENT,
              border: "none",
              borderRadius: 7,
              padding: "8px 16px",
              cursor: "pointer",
            }}
          >
            돌아가기
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}
// ================= MAIN APP =================
const CHECK_STATUS_FILTERS = [
  {
    id: "all",
    label: "전체",
    test: () => true,
  },
  {
    id: "fail",
    label: "확인 필요",
    test: (r) => !r.pass,
  },
  {
    id: "pass",
    label: "충족",
    test: (r) => r.pass,
  },
  {
    id: "missing",
    label: "신청 누락",
    test: (r) => r.missingSemesters.length > 0,
  },
  {
    id: "transfer",
    label: "전입생",
    test: (r) => r.isTransfer,
  },
  {
    id: "changed",
    label: "상담·변경",
    test: (r) => r.hasChanges,
  },
  {
    id: "unapplied",
    label: "변경 미반영",
    test: (r) => r.change.applyState === "pending" || r.change.applyState === "stale",
  },
  {
    id: "extra",
    label: "공동교육과정 등",
    test: (r) => r.extraEntries.length > 0,
  },
];
function GraduationCheckerInner() {
  const [page, setPage] = useState("curriculum"); // curriculum | builder | roster | check | change | transfer | extra
  // ---- curriculum state ----
  const [curriculumFile, setCurriculumFile] = useState(null); // {name, origin, courses, groups, subjectMaster}
  const [curriculumError, setCurriculumError] = useState(null);
  const [minCredit, setMinCredit] = useState(174);
  const [reqOverride, setReqOverride] = useState({});
  const [showReq, setShowReq] = useState(false);
  const [priorityGroups, setPriorityGroups] = useState({
    "예술": true,
    "생활교양": true,
  }); // group -> bool ('중점 관리 교과군')
  const [admissionYear, setAdmissionYear] = useState(new Date().getFullYear());
  const [creativeActivityCredit, setCreativeActivityCredit] = useState(18);
  // ---- builder ('편제표 직접 작성') state ----
  const [builderData, setBuilderData] = useState(emptyBuilderData);
  // ---- registration state ----
  const [regFiles, setRegFiles] = useState([]); // {id, source:'hakjeom'|'apin', name, label, semesterMode, status:'ok'|'error'|'wrongZone', detected, error, count, records}
  const [semesterSourceChoice, setSemesterSourceChoice] = useState({}); // semester code -> file id, when several files cover the same semester
  const [mergeGroups, setMergeGroups] = useState([]); // {id, recordIds: [uid,...]}
  const [checked, setChecked] = useState({}); // identity-unit key -> bool, scratch selection for merge UI
  const [classMappingFiles, setClassMappingFiles] = useState([]); // {name, status, error, entries}
  // ---- check page state ----
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [selected, setSelected] = useState(null);
  const [selectedClass, setSelectedClass] = useState(null);
  const [archivedStudents, setArchivedStudents] = useState({}); // key -> { reason, archivedAt, name }
  const [archiveModalKey, setArchiveModalKey] = useState(null);
  const [archiveReason, setArchiveReason] = useState("");
  const [classTeachers, setClassTeachers] = useState({}); // classLabel -> teacher name
  const [printKeys, setPrintKeys] = useState(null); // array of result keys queued for printing, or null
  // ---- per-student adjustments (④ 선택과목 변경 · ⑤ 전입생 · ⑥ 공동교육과정) ----
  const [courseChanges, setCourseChanges] = useState({}); // studentKey -> {name, removed:[entryKey], added:[{name, semester}], memo, updatedAt}
  const [neisFiles, setNeisFiles] = useState([]); // 나이스 교과학습발달상황 엑셀 {id, name, status, error, classInfo, students, warnings} — 저장 파일에는 넣지 않습니다
  const [transferInfo, setTransferInfo] = useState({}); // studentKey -> {name, firstSem, prevSchool, memo, priorCourses:[{id, semester, group, name, credit}]}
  const [extraCourses, setExtraCourses] = useState([]); // [{id, kind, name, group, credit, semester, memo, studentKeys, studentNames}]
  const [manualStudents, setManualStudents] = useState([]); // [{id, name, grade, classNum, number}] — students not in any registration file
  // ---- ② 학생 기초조사 ----
  const [rosterMode, setRosterMode] = useState("upload"); // 'upload' (수강신청 엑셀) | 'survey' (학생 기초조사)
  const [surveyConfig, setSurveyConfig] = useState(emptySurveyConfig);
  const [surveyFiles, setSurveyFiles] = useState([]); // {id, name, status, error, meta, skipped, submissions}
  const [surveyChoice, setSurveyChoice] = useState({}); // 'grade-class-number' -> submission key, when one number has several names
  const [surveyExcluded, setSurveyExcluded] = useState({}); // submission key -> true
  const [surveyPreview, setSurveyPreview] = useState(false);
  const [grade1Electives, setGrade1Electives] = useState(null); // 'none' | 'some' | null — 1학년 1·2학기에 학생선택 과목이 있는지 선생님이 확인한 답
  const [printQr, setPrintQr] = useState(null); // {title, grade, deadline, items:[{c, url}]} while printing QR 안내문
  const [printList, setPrintList] = useState(null); // {title, subtitle, columns, rows} while printing a ③ 학생 목록
  const [changeKey, setChangeKey] = useState(null);
  const [transferKey, setTransferKey] = useState(null);
  // ---- save/load state ----
  const [saveModal, setSaveModal] = useState(false);
  const [usePassword, setUsePassword] = useState(false);
  const [savePassword, setSavePassword] = useState("");
  const [saveError, setSaveError] = useState(null);
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [pendingLoad, setPendingLoad] = useState(null); // encrypted payload waiting for password
  const [loadPassword, setLoadPassword] = useState("");
  const [loadError, setLoadError] = useState(null);
  const loadInputRef = useRef(null);
  const cryptoAvailable = typeof window !== "undefined" && !!(window.crypto && window.crypto.subtle);
  // ---- 브라우저 임시 보관 (비밀번호로 잠가 IndexedDB에 자동 저장) ----
  const [keepPassword, setKeepPassword] = useState(null); // 정해지면 자동 보관이 켜진 것. 메모리에만 둡니다
  const [keepDraft, setKeepDraft] = useState(""); // 보관 띠의 비밀번호 입력칸
  const [keepDeclined, setKeepDeclined] = useState(() => {
    try {
      return sessionStorage.getItem("keepDeclined") === "1";
    } catch (e) {
      return false;
    }
  });
  const [keepState, setKeepState] = useState({ savedAt: null, error: "" });
  const [resumeInfo, setResumeInfo] = useState(null); // 다시 열었을 때 발견한 지난 작업 {savedAt, summary, payload}
  const [resumePassword, setResumePassword] = useState("");
  const [resumeError, setResumeError] = useState("");
  const [resumeBusy, setResumeBusy] = useState(false);
  const keeperRef = useRef(null);
  const keepDirtyRef = useRef(false);
  const keepTimerRef = useRef(null);
  useEffect(() => {
    // 처음 열 때: 이 브라우저에 지난 작업이 남아 있으면 이어할지 묻습니다 (묻지 않고 자동으로 불러오지 않음)
    if (!cryptoAvailable) return;
    keepRead()
      .then((rec) => {
        if (rec) setResumeInfo(rec);
      })
      .catch(() => {});
  }, []);
  const startKeeping = async (password) => {
    keeperRef.current = await makeKeeper(password);
    setKeepPassword(password);
    setKeepDraft("");
    keepDirtyRef.current = true; // 바로 한 번 보관
  };
  const stopKeeping = async () => {
    keeperRef.current = null;
    setKeepPassword(null);
    setKeepState({ savedAt: null, error: "" });
    await keepClear();
  };
  const declineKeeping = () => {
    setKeepDeclined(true);
    try {
      sessionStorage.setItem("keepDeclined", "1");
    } catch (e) {
      // 세션 저장이 막혀 있으면 이번 화면에서만 기억
    }
  };
  const handleCurriculum = async (files) => {
    const file = files[0];
    setCurriculumError(null);
    try {
      const wb = await readWorkbookFromFile(file);
      const parsedCourses = parseCourseList(wb);
      if (parsedCourses.length === 0)
        throw new Error(
          "편제표에서 과목을 하나도 찾지 못했습니다. ‘편제표입력’ 시트의 머리글(구분·교과(군)·과목구분·과목명…)을 확인해주세요.",
        );
      const { req: parsedReq, groups: parsedGroups } = parseGroupRequirements(wb, parsedCourses);
      const { courses, req, groups } = mergeLifestyleGroups(parsedCourses, parsedReq, parsedGroups);
      const subjectMasterRaw = parseSubjectMaster(wb);
      const subjectMaster = new Map(
        Array.from(subjectMasterRaw.entries()).map(([k, v]) => [
          k,
          {
            ...v,
            group: toCurriculumGroup(v.group),
          },
        ]),
      );
      setCurriculumFile({
        name: file.name,
        origin: "upload",
        courses,
        groups,
        subjectMaster,
      });
      setReqOverride(req);
      setBuilderData(coursesToBuilderData(parsedCourses));
    } catch (e) {
      setCurriculumError(e.message || "편제표 파일을 읽을 수 없습니다.");
      setCurriculumFile(null);
    }
  };
  const applyBuilderData = () => {
    const rawCourses = builderDataToCourses(builderData);
    if (rawCourses.length === 0) {
      window.alert("편제표에 담긴 과목이 없습니다. 학기별로 학교지정 과목이나 선택그룹 과목을 먼저 담아주세요.");
      return;
    }
    const { req: parsedReq, groups: parsedGroups } = nationalOnlyRequirements(rawCourses);
    const { courses, req, groups } = mergeLifestyleGroups(rawCourses, parsedReq, parsedGroups);
    const subjectMaster = new Map(
      Array.from(buildNationalSubjectMasterMap().entries()).map(([k, v]) => [
        k,
        {
          ...v,
          group: toCurriculumGroup(v.group),
        },
      ]),
    );
    setCurriculumError(null);
    setCurriculumFile({
      name: "직접 작성한 편제표 (수작업)",
      origin: "builder",
      courses,
      groups,
      subjectMaster,
    });
    setReqOverride(req);
    setPage("roster");
  };
  const handleRegFiles = async (files, source) => {
    const parsed = [];
    for (const file of files) {
      const base = {
        id: newId("rf"),
        source,
        name: file.name,
        label: file.name.replace(/(\.(xlsx|xls))+$/i, ""),
        addedAt: Date.now(),
        // 같은 학기 자료가 여러 곳에 있을 때 '나중에 올린 자료'를 기본으로 고르는 데 씁니다
        semesterMode: "auto", // 'auto' | a semester code chosen from the dropdown (only used when the file itself carries no semester)
      };
      try {
        const wb = await readWorkbookFromFile(file);
        const detected = detectRegistrationFormat(wb);
        if (!detected) {
          parsed.push({
            ...base,
            status: "error",
            error:
              "수강신청 파일 형식을 인식하지 못했습니다. 고교학점제 누리집이나 압핀에서 내려받은 원본 파일인지 확인해주세요.",
            records: [],
          });
        } else if (detected.source !== source) {
          // parsed with the right parser already, but kept out of the calculation until the user moves it
          parsed.push({
            ...base,
            status: "wrongZone",
            detected: detected.source,
            count: detected.records.length,
            records: detected.records,
          });
        } else {
          parsed.push({
            ...base,
            status: "ok",
            count: detected.records.length,
            records: detected.records,
          });
        }
      } catch (e) {
        parsed.push({
          ...base,
          status: "error",
          error: e.message || "읽기 실패",
          records: [],
        });
      }
    }
    setRegFiles((prev) => [...prev, ...parsed]);
  };
  const removeRegFile = (id) => setRegFiles((prev) => prev.filter((f) => f.id !== id));
  const relabelRegFile = (id, label) =>
    setRegFiles((prev) =>
      prev.map((f) =>
        f.id === id
          ? {
              ...f,
              label,
            }
          : f,
      ),
    );
  const setRegFileSemester = (id, semesterMode) =>
    setRegFiles((prev) =>
      prev.map((f) =>
        f.id === id
          ? {
              ...f,
              semesterMode,
            }
          : f,
      ),
    );
  const moveRegFile = (id) =>
    setRegFiles((prev) =>
      prev.map((f) =>
        f.id === id && f.status === "wrongZone"
          ? {
              ...f,
              source: f.detected,
              status: "ok",
              detected: undefined,
            }
          : f,
      ),
    );
  const setSemesterSource = (sem, fileId) =>
    setSemesterSourceChoice((prev) => ({
      ...prev,
      [sem]: fileId,
    }));
  const toggleRegFileDisabled = (id) =>
    setRegFiles((prev) =>
      prev.map((f) =>
        f.id === id
          ? {
              ...f,
              disabled: !f.disabled,
            }
          : f,
      ),
    );
  // ---- 학생 기초조사: 조사 내용(편제표로 만듦) · 응답 → 학생 기록 ----
  const semesterPoolDefs = useMemo(
    () => (curriculumFile ? buildSemesterPoolDefs(curriculumFile.courses) : []),
    [curriculumFile],
  );
  const surveyPayload = useMemo(
    () =>
      buildSurveyPayload(surveyConfig, semesterPoolDefs, {
        admissionYear,
      }),
    [surveyConfig, semesterPoolDefs, admissionYear],
  );
  const surveyIssues = useMemo(
    () => surveyProblems(surveyPayload, surveyConfig, !!curriculumFile),
    [surveyPayload, surveyConfig, curriculumFile],
  );
  const surveyData = useMemo(
    () =>
      buildSurveyData(surveyFiles, {
        choice: surveyChoice,
        excluded: surveyExcluded,
        classes: surveyPayload.classes,
        grade: surveyPayload.grade,
      }),
    [surveyFiles, surveyChoice, surveyExcluded, surveyPayload],
  );
  const surveyPreviewHtml = useMemo(
    () =>
      surveyPreview
        ? buildStudentPageHtml({
            survey: surveyPayload,
            ban: surveyPayload.classes.length ? surveyPayload.classes[0].c : null,
            closed: false,
            preview: true,
          })
        : "",
    [surveyPreview, surveyPayload],
  );
  const handleSurveyFiles = async (files) => {
    const parsed = [];
    for (const file of files) {
      const base = {
        id: newId("sf"),
        name: file.name,
        addedAt: Date.now(),
        disabled: false,
      };
      try {
        const wb = await readWorkbookFromFile(file);
        const r = parseSurveyWorkbook(wb);
        parsed.push({
          ...base,
          status: "ok",
          error: "",
          meta: r.meta,
          skipped: r.skipped,
          submissions: r.submissions,
        });
      } catch (e) {
        parsed.push({
          ...base,
          status: "error",
          error: e.message || "읽기 실패",
          meta: null,
          skipped: 0,
          submissions: [],
        });
      }
    }
    setSurveyFiles((prev) => [...prev, ...parsed]);
  };
  // [구글 시트에서 지금 가져오기]: 배포한 웹 앱 주소에 열쇠를 붙여 응답을 받아옵니다. 가져온 것은 파일 하나처럼 다루되,
  // 다시 가져오면 이전 것과 바꿉니다 (조사 중간에 여러 번 눌러도 목록이 늘어나지 않게)
  const [surveyFetch, setSurveyFetch] = useState({ busy: false, error: "" });
  const fetchSurveyLive = async () => {
    const url = checkWebAppUrl(surveyConfig.webAppUrl);
    if (!url.ok) return;
    setSurveyFetch({ busy: true, error: "" });
    const exportUrl = `${url.url}?export=${encodeURIComponent(surveyConfig.exportKey)}`;
    const OLD_CODE_MSG =
      "시트가 응답 대신 학생 화면을 돌려줍니다. 응답 내보내기가 없는 옛 코드로 배포된 시트입니다. 시트의 Apps Script에 최신 코드를 붙여넣고 [배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포]를 해 주세요.";
    try {
      let json;
      try {
        const res = await fetch(exportUrl, { redirect: "follow" });
        if (!res.ok) throw new Error(`주소에 연결하지 못했습니다 (${res.status}).`);
        const text = await res.text();
        if (/^\s*</.test(text)) throw new Error(OLD_CODE_MSG);
        json = JSON.parse(text);
      } catch (e) {
        // 브라우저가 직접 요청을 막았으면(주로 CORS) 스크립트 태그 방식으로 한 번 더
        if (!(e instanceof TypeError)) throw e;
        json = await loadJsonp(exportUrl);
      }
      if (json && json.error) throw new Error(json.error);
      if (!json || json.kind !== "graduation-survey-export" || !json.sheets)
        throw new Error(
          "응답 내보내기를 지원하지 않는 시트입니다. 시트 코드를 최신으로 바꾸고 [배포 관리 → 새 버전]으로 다시 배포해 주세요.",
        );
      const wb = XLSX.utils.book_new();
      Object.keys(json.sheets).forEach((name) => {
        XLSX.utils.book_append_sheet(
          wb,
          XLSX.utils.aoa_to_sheet(Array.isArray(json.sheets[name]) ? json.sheets[name] : []),
          name,
        );
      });
      const r = parseSurveyWorkbook(wb);
      const at = String(json.at || "").slice(0, 16);
      const entry = {
        id: "sf-live",
        name: `구글 시트에서 가져옴 (${at || formatDateYmd()})`,
        addedAt: Date.now(),
        disabled: false,
        live: true,
        status: "ok",
        error: "",
        meta: r.meta,
        skipped: r.skipped,
        submissions: r.submissions,
      };
      setSurveyFiles((prev) => [...prev.filter((f) => f.id !== "sf-live"), entry]);
      setSurveyFetch({ busy: false, error: "" });
    } catch (e) {
      const msg = e && e.message ? e.message : String(e);
      setSurveyFetch({
        busy: false,
        error: /fetch|network|Failed|JSONP/i.test(msg)
          ? "시트에 연결하지 못했습니다. 인터넷 연결과 5번 칸의 주소를 확인하고, 옆의 [주소 열어 확인]으로 시트가 무엇을 돌려주는지 보세요. 학생 화면이 보이면 옛 코드로 배포된 것이니 새 버전으로 다시 배포해야 합니다. 계속 안 되면 시트에서 엑셀로 내려받아 올려 주세요."
          : msg,
      });
    }
  };
  const removeSurveyFile = (id) => setSurveyFiles((prev) => prev.filter((f) => f.id !== id));
  const toggleSurveyFileDisabled = (id) =>
    setSurveyFiles((prev) =>
      prev.map((f) =>
        f.id === id
          ? {
              ...f,
              disabled: !f.disabled,
            }
          : f,
      ),
    );
  const setSurveyChoiceFor = (identity, key) =>
    setSurveyChoice((prev) => ({
      ...prev,
      [identity]: key,
    }));
  const toggleSurveyExcluded = (key) =>
    setSurveyExcluded((prev) => {
      const next = {
        ...prev,
      };
      if (next[key]) delete next[key];
      else next[key] = true;
      return next;
    });
  // ---- 학기마다 자료 하나: 수강신청 파일 또는 학생 기초조사 (see the pure helpers above) ----
  const regFileInfos = useMemo(() => {
    const courses = (curriculumFile === null || curriculumFile === void 0 ? void 0 : curriculumFile.courses) || [];
    return new Map(regFiles.map((f) => [f.id, computeRegFileInfo(f, courses)]));
  }, [regFiles, curriculumFile]);
  const surveyProvider = useMemo(
    () =>
      surveyData.records.length > 0
        ? {
            semesters: Array.from(surveyData.coveredSemesters),
            addedAt: surveyData.lastAddedAt,
          }
        : null,
    [surveyData],
  );
  const semesterProviders = useMemo(
    () => buildSemesterProviders(regFiles, regFileInfos, surveyProvider),
    [regFiles, regFileInfos, surveyProvider],
  );
  const chosenSourceBySem = useMemo(
    () => pickSourceBySem(semesterProviders, semesterSourceChoice),
    [semesterProviders, semesterSourceChoice],
  );
  const semesterConflicts = useMemo(
    () =>
      Array.from(semesterProviders.entries())
        .filter(([, ids]) => ids.length > 1)
        .map(([sem, fileIds]) => ({
          sem,
          fileIds,
          chosenId: chosenSourceBySem.get(sem),
          unresolved: !fileIds.includes(semesterSourceChoice[sem]),
        }))
        .sort((a, b) => compareSemester(a.sem, b.sem)),
    [semesterProviders, chosenSourceBySem, semesterSourceChoice],
  );
  const unresolvedSourceSems = useMemo(
    () => semesterConflicts.filter((c) => c.unresolved).map((c) => c.sem),
    [semesterConflicts],
  );
  const fileRawRecords = useMemo(
    () => buildRawRecords(regFiles, regFileInfos, chosenSourceBySem),
    [regFiles, regFileInfos, chosenSourceBySem],
  );
  const surveyRawRecords = useMemo(
    () => buildSurveyRawRecords(surveyData, chosenSourceBySem),
    [surveyData, chosenSourceBySem],
  );
  const rawRecords = useMemo(() => [...fileRawRecords, ...surveyRawRecords], [fileRawRecords, surveyRawRecords]);
  const usesSurvey = surveyRawRecords.length > 0;
  // 이름이 겹치는 학생: 수강신청 파일 안에서 겹치거나, 기초조사에서 서로 다른 번호로 같은 이름이 낸 경우
  const ambiguousNames = useMemo(() => {
    const names = new Set(computeAmbiguousNames(regFiles));
    if (usesSurvey) surveyData.ambiguousNames.forEach((n) => names.add(n));
    return names;
  }, [regFiles, usesSurvey, surveyData]);
  const sourceSemesters = useMemo(() => {
    const upload = [];
    const survey = [];
    chosenSourceBySem.forEach((id, sem) => (id === SURVEY_SOURCE_ID ? survey : upload).push(sem));
    return {
      upload: upload.sort(compareSemester),
      survey: survey.sort(compareSemester),
    };
  }, [chosenSourceBySem]);
  const grade1Status = useMemo(
    () => grade1ElectiveStatus(semesterPoolDefs, grade1Electives),
    [semesterPoolDefs, grade1Electives],
  );
  const uidToMergeGroup = useMemo(() => {
    const m = new Map();
    mergeGroups.forEach((g) => g.recordIds.forEach((uid) => m.set(uid, g.id)));
    return m;
  }, [mergeGroups]);
  const fileMappingEntries = useMemo(
    () => classMappingFiles.filter((f) => f.status === "ok").flatMap((f) => f.entries),
    [classMappingFiles],
  );
  const hakjeomEntries = useMemo(() => hakjeomMappingEntries(regFiles), [regFiles]);
  const classMappingIndex = useMemo(
    () => buildClassMappingIndex([...fileMappingEntries, ...hakjeomEntries]),
    [fileMappingEntries, hakjeomEntries],
  );
  const fileMappingIndex = useMemo(() => buildClassMappingIndex(fileMappingEntries), [fileMappingEntries]);
  const hakjeomMappingIndex = useMemo(() => buildClassMappingIndex(hakjeomEntries), [hakjeomEntries]);
  // identity of a same-name record before any manual merge: through the class-mapping file when possible, otherwise
  // (name + 학년-반-번호). The 압핀 2-1/2-2 files share one roster, so the same student carries the same 학년-반-번호 in
  // both and is linked automatically; only across a grade change (1학년 반 → 2학년 학번) is a manual merge needed.
  const { identityKeyFor, groupKeyFor } = makeStudentKeyFns(classMappingIndex, ambiguousNames, uidToMergeGroup);
  const studentGroups = useMemo(
    () => buildStudentGroups(rawRecords, groupKeyFor, manualStudents),
    [rawRecords, ambiguousNames, uidToMergeGroup, classMappingIndex, manualStudents],
  );
  const isMappedRecord = (r, index = fileMappingIndex) => {
    const k = idKeyOf(r.name, r.idFields);
    return !!(k && index.has(k));
  };
  // counted per identity unit (records auto-linked by 학년-반-번호 count once), excluding merged / mapping-linked ones
  const unresolvedCount = useMemo(() => {
    const keys = new Set();
    rawRecords.forEach((r) => {
      if (!ambiguousNames.has(r.name) || uidToMergeGroup.has(r.uid)) return;
      const k = identityKeyFor(r);
      if (!k.startsWith("mapped:")) keys.add(k);
    });
    return keys.size;
  }, [rawRecords, ambiguousNames, uidToMergeGroup, classMappingIndex]);
  const mappedResolvedCount = useMemo(() => {
    const names = new Set();
    rawRecords.forEach((r) => {
      if (ambiguousNames.has(r.name) && isMappedRecord(r)) names.add(r.name);
    });
    return names.size;
  }, [rawRecords, ambiguousNames, fileMappingIndex]);
  // 누리집 파일의 1학년 학번으로 연결된 동명이인(편성 자료로 이미 연결된 이름은 빼고 셉니다)
  const hakjeomLinkedCount = useMemo(() => {
    const names = new Set();
    rawRecords.forEach((r) => {
      if (ambiguousNames.has(r.name) && isMappedRecord(r, hakjeomMappingIndex) && !isMappedRecord(r)) names.add(r.name);
    });
    return names.size;
  }, [rawRecords, ambiguousNames, hakjeomMappingIndex, fileMappingIndex]);
  const handleClassMappingFiles = async (files) => {
    const parsed = [];
    for (const file of files) {
      try {
        const wb = await readWorkbookFromFile(file);
        const entries = parseClassMappingFile(wb);
        parsed.push({
          name: file.name,
          status: "ok",
          entries,
          count: entries.length,
        });
      } catch (e) {
        parsed.push({
          name: file.name,
          status: "error",
          error: e.message || "읽기 실패",
          entries: [],
        });
      }
    }
    setClassMappingFiles((prev) => [...prev, ...parsed]);
  };
  const removeClassMappingFile = (idx) => setClassMappingFiles((prev) => prev.filter((_, i) => i !== idx));
  // ---- merge UI actions ----
  const toggleCheck = (unitKey) =>
    setChecked((prev) => ({
      ...prev,
      [unitKey]: !prev[unitKey],
    }));
  // uids: every record of the checked identity units · unitKeys: the checkbox keys to clear afterwards
  const mergeChecked = (uids, unitKeys) => {
    if (uids.length < 2 || unitKeys.length < 2) return;
    setMergeGroups((prev) => [
      ...prev,
      {
        id: `merge-${prev.length}-${Date.now()}`,
        recordIds: uids,
      },
    ]);
    setChecked((prev) => {
      const next = {
        ...prev,
      };
      unitKeys.forEach((k) => delete next[k]);
      return next;
    });
  };
  const unmerge = (groupId) => setMergeGroups((prev) => prev.filter((g) => g.id !== groupId));
  // ---- credit computation ----
  const courseMap = useMemo(
    () =>
      buildCourseMap((curriculumFile === null || curriculumFile === void 0 ? void 0 : curriculumFile.courses) || []),
    [curriculumFile],
  );
  const groups = useMemo(
    () => (curriculumFile === null || curriculumFile === void 0 ? void 0 : curriculumFile.groups) || [],
    [curriculumFile],
  );
  const orderedGroups = useMemo(() => {
    return [...groups].sort((a, b) => {
      const pa = priorityGroups[a] ? 0 : 1;
      const pb = priorityGroups[b] ? 0 : 1;
      return pa - pb;
    });
  }, [groups, priorityGroups]);
  const curriculumIssues = useMemo(
    () => (curriculumFile ? validateCurriculum(curriculumFile.courses) : []),
    [curriculumFile],
  );
  const semesterSummary = useMemo(() => {
    if (!curriculumFile) return [];
    const allTakenSubjectKeys = new Set(rawRecords.flatMap((r) => r.subjects.map((s) => entryKey(s.name, s.semester))));
    const bySem = new Map();
    curriculumFile.courses.forEach((c) => {
      const sem = c.semester || "학기 미상";
      if (!bySem.has(sem))
        bySem.set(sem, {
          mandatory: [],
          elective: [],
        });
      const bucket = bySem.get(sem);
      if (c.gubun === "학교지정") bucket.mandatory.push(c);
      else bucket.elective.push(c);
    });
    return Array.from(bySem.entries())
      .map(([sem, { mandatory, elective }]) => {
        const isAuto = elective.length === 0;
        const coveredCount = elective.filter((c) => allTakenSubjectKeys.has(`${c.key}||${c.semester}`)).length;
        let status;
        if (isAuto) status = "auto";
        else if (coveredCount === 0) status = "missing";
        else if (coveredCount === elective.length) status = "full";
        else status = "partial";
        return {
          sem,
          label: formatSemester(sem),
          mandatory,
          elective,
          isAuto,
          coveredCount,
          status,
        };
      })
      .sort((a, b) => compareSemester(a.sem, b.sem));
  }, [curriculumFile, rawRecords]);
  const latestGradeInData = useMemo(() => computeLatestGrade(studentGroups), [studentGroups]);
  const coveredSemesters = useMemo(() => new Set(semesterProviders.keys()), [semesterProviders]);
  // ③ 학급별 표에 보여줄 기초조사 미제출 번호 ('2학년 3반' → [5, 17])
  const surveyMissingByClass = useMemo(() => {
    if (!usesSurvey) return null;
    return new Map(surveyData.missingByClass.map((k) => [`${surveyPayload.grade}학년 ${k.c}반`, k.missing]));
  }, [usesSurvey, surveyData, surveyPayload]);
  const results = useMemo(() => {
    if (!curriculumFile) return [];
    return computeAllResults(
      studentGroups,
      {
        courses: curriculumFile.courses,
        groups,
        subjectMaster: curriculumFile.subjectMaster,
        courseMap,
        poolDefs: semesterPoolDefs,
        reqOverride,
        minCredit,
        latestGradeInData,
        coveredSemesters,
      },
      {
        courseChanges,
        transferInfo,
        extraCourses,
      },
    );
  }, [
    studentGroups,
    curriculumFile,
    groups,
    courseMap,
    semesterPoolDefs,
    reqOverride,
    minCredit,
    latestGradeInData,
    coveredSemesters,
    courseChanges,
    transferInfo,
    extraCourses,
  ]);
  const resultByKey = useMemo(() => new Map(results.map((r) => [r.key, r])), [results]);
  const activeResults = useMemo(() => results.filter((r) => !archivedStudents[r.key]), [results, archivedStudents]);
  const archivedResults = useMemo(
    () =>
      results
        .filter((r) => archivedStudents[r.key])
        .map((r) => ({
          ...r,
          archiveInfo: archivedStudents[r.key],
        })),
    [results, archivedStudents],
  );
  const classSummary = useMemo(() => {
    const m = new Map();
    activeResults.forEach((r) => {
      const label = r.classLabel;
      if (!m.has(label))
        m.set(label, {
          label,
          sortKey: r.classSortKey,
          unclassified: r.isUnclassified,
          count: 0,
          passCount: 0,
          failCount: 0,
          missingCount: 0,
          changedCount: 0,
          unappliedCount: 0,
          transferCount: 0,
        });
      const c = m.get(label);
      c.count += 1;
      if (r.pass) c.passCount += 1;
      else c.failCount += 1;
      if (r.missingSemesters.length > 0) c.missingCount += 1;
      if (r.hasChanges) c.changedCount += 1;
      if (r.change.applyState === "pending" || r.change.applyState === "stale") c.unappliedCount += 1;
      if (r.isTransfer) c.transferCount += 1;
    });
    return Array.from(m.values()).sort((a, b) => a.sortKey - b.sortKey || a.label.localeCompare(b.label, "ko"));
  }, [activeResults]);
  const checkFilters = useMemo(
    () =>
      CHECK_STATUS_FILTERS.map((f) => ({
        ...f,
        count: activeResults.filter(f.test).length,
      })),
    [activeResults],
  );
  const filtered = useMemo(() => {
    const f = CHECK_STATUS_FILTERS.find((x) => x.id === statusFilter) || CHECK_STATUS_FILTERS[0];
    let r = activeResults.filter(f.test);
    if (selectedClass) {
      r = r.filter((x) => x.classLabel === selectedClass);
    } else {
      // '전체' view: unclassified (likely transferred/dropped) students float to the top
      r = [...r].sort((a, b) => (a.isUnclassified === b.isUnclassified ? 0 : a.isUnclassified ? -1 : 1));
    }
    const q = search.trim();
    if (q) r = r.filter((x) => x.name.includes(q) || x.studentId.includes(q));
    return r;
  }, [activeResults, search, selectedClass, statusFilter]);
  const passCount = activeResults.filter((r) => r.pass).length;
  const failCount = activeResults.length - passCount;
  const selectedResult = selected ? resultByKey.get(selected) || null : null;
  const printResults = useMemo(() => {
    if (!printKeys) return [];
    return printKeys.map((k) => resultByKey.get(k)).filter(Boolean);
  }, [printKeys, resultByKey]);
  const printOne = (key) => {
    setPrintQr(null);
    setPrintList(null);
    setPrintKeys([key]);
    setTimeout(() => window.print(), 80);
  };
  const printClass = (classLabel, onlyFail) => {
    const keys = activeResults
      .filter((r) => (classLabel === null ? true : r.classLabel === classLabel))
      .filter((r) => (onlyFail ? !r.pass : true))
      .sort(
        (a, b) =>
          a.classSortKey - b.classSortKey ||
          a.studentId.localeCompare(b.studentId, "ko", {
            numeric: true,
          }),
      )
      .map((r) => r.key);
    if (keys.length === 0) {
      window.alert("인쇄할 학생이 없습니다.");
      return;
    }
    setPrintQr(null);
    setPrintList(null);
    setPrintKeys(keys);
    setTimeout(() => window.print(), 80);
  };
  const printProblemList = (sheet) => {
    setPrintKeys(null);
    setPrintQr(null);
    setPrintList(sheet);
    setTimeout(() => window.print(), 120);
  };
  const printSurveyQr = (baseUrl) => {
    setPrintKeys(null);
    setPrintList(null);
    setPrintQr({
      title: surveyPayload.title,
      grade: surveyPayload.grade,
      deadline: surveyPayload.deadline,
      items: surveyPayload.classes.map((k) => ({
        c: k.c,
        url: surveyClassUrl(baseUrl, k.c),
      })),
    });
    setTimeout(() => window.print(), 150);
  };
  const confirmArchive = () => {
    if (!archiveModalKey) return;
    const r = resultByKey.get(archiveModalKey);
    setArchivedStudents((prev) => ({
      ...prev,
      [archiveModalKey]: {
        reason: archiveReason.trim(),
        archivedAt: new Date().toISOString(),
        name: r ? r.name : nameFromStudentKey(archiveModalKey),
      },
    }));
    setArchiveModalKey(null);
    setArchiveReason("");
    if (selected === archiveModalKey) setSelected(null);
  };
  const restoreStudent = (key) => {
    setArchivedStudents((prev) => {
      const next = {
        ...prev,
      };
      delete next[key];
      return next;
    });
  };
  // ---- ④ 선택과목 변경 ----
  const updateCourseChange = (studentKey, studentName, updater) =>
    setCourseChanges((prev) => {
      const cur = prev[studentKey] || {
        name: studentName,
        removed: [],
        added: [],
        memo: "",
        updatedAt: null,
      };
      const next = {
        ...updater(cur),
        name: studentName || cur.name,
        updatedAt: new Date().toISOString(),
      };
      const copy = {
        ...prev,
      };
      if (next.removed.length === 0 && next.added.length === 0 && !next.memo) delete copy[studentKey];
      else copy[studentKey] = next;
      return copy;
    });
  const toggleCourseChange = (result, name, semester) => {
    const k = entryKey(name, semester);
    const isOriginal = result.registered.some((e) => entryKey(e.name, e.semester) === k);
    updateCourseChange(result.key, result.name, (cur) => {
      if (isOriginal) {
        return {
          ...cur,
          removed: cur.removed.includes(k) ? cur.removed.filter((x) => x !== k) : [...cur.removed, k],
        };
      }
      const exists = cur.added.some((a) => entryKey(a.name, a.semester) === k);
      return {
        ...cur,
        added: exists
          ? cur.added.filter((a) => entryKey(a.name, a.semester) !== k)
          : [
              ...cur.added,
              {
                name,
                semester,
              },
            ],
      };
    });
  };
  const setChangeMemo = (result, memo) =>
    updateCourseChange(result.key, result.name, (cur) => ({
      ...cur,
      memo,
    }));
  const setChangeApplied = (result, on) =>
    setCourseChanges((prev) => {
      const cur = prev[result.key];
      if (!cur) return prev;
      return {
        ...prev,
        [result.key]: {
          ...cur,
          applied: on
            ? {
                at: new Date().toISOString(),
                sig: changeSignature(cur),
              }
            : null,
        },
      };
    });
  const resetCourseChange = (key) =>
    setCourseChanges((prev) => {
      const copy = {
        ...prev,
      };
      delete copy[key];
      return copy;
    });
  // ---- ⑤ 전입생 ----
  const registerTransfer = (result) => {
    setTransferInfo((prev) => {
      if (prev[result.key]) return prev;
      const firstRegistered = (result.registered || [])
        .map((e) => e.semester)
        .filter((s) => SEMESTER_CODES.includes(s))
        .sort(compareSemester)[0];
      const suggested = firstRegistered && SEMESTER_CODES.indexOf(firstRegistered) > 0 ? firstRegistered : "2-1";
      return {
        ...prev,
        [result.key]: {
          name: result.name,
          firstSem: suggested,
          prevSchool: "",
          memo: "",
          priorCourses: [],
        },
      };
    });
    setTransferKey(result.key);
  };
  const handleNeisFiles = async (files) => {
    const added = [];
    for (const file of files) {
      try {
        const wb = await readWorkbookFromFile(file);
        const parsed = parseNeisGradesWorkbook(wb);
        added.push({
          id: newId("neis"),
          name: file.name,
          status: "ok",
          error: "",
          ...parsed,
        });
      } catch (e) {
        added.push({
          id: newId("neis"),
          name: file.name,
          status: "error",
          error: e.message || String(e),
          classInfo: null,
          students: [],
          warnings: [],
        });
      }
    }
    setNeisFiles((prev) => [...prev, ...added]);
  };
  const removeNeisFile = (id) => setNeisFiles((prev) => prev.filter((f) => f.id !== id));
  // 나이스에서 가져온 과목으로 그 학기들의 전적교 과목을 바꿉니다(다른 학기 입력은 그대로)
  const applyNeisRows = (key, rows, semesters) =>
    updateTransfer(key, (cur) => ({
      priorCourses: [...(cur.priorCourses || []).filter((p) => !semesters.includes(p.semester)), ...rows],
    }));
  const updateTransfer = (key, patch) =>
    setTransferInfo((prev) =>
      prev[key]
        ? {
            ...prev,
            [key]: {
              ...prev[key],
              ...(typeof patch === "function" ? patch(prev[key]) : patch),
            },
          }
        : prev,
    );
  const removeTransfer = (key) => {
    setTransferInfo((prev) => {
      const copy = {
        ...prev,
      };
      delete copy[key];
      return copy;
    });
    setTransferKey((k) => (k === key ? null : k));
  };
  const addManualStudent = ({ name, grade, classNum, number }) => {
    const id = newId("ms");
    const key = "manual:" + id;
    const trimmed = String(name || "").trim();
    setManualStudents((prev) => [
      ...prev,
      {
        id,
        name: trimmed,
        grade,
        classNum,
        number,
      },
    ]);
    setTransferInfo((prev) => ({
      ...prev,
      [key]: {
        name: trimmed,
        firstSem: "2-1",
        prevSchool: "",
        memo: "",
        priorCourses: [],
      },
    }));
    setTransferKey(key);
  };
  const updateManualStudent = (id, patch) =>
    setManualStudents((prev) =>
      prev.map((m) =>
        m.id === id
          ? {
              ...m,
              ...patch,
            }
          : m,
      ),
    );
  const removeManualStudent = (id) => {
    const key = "manual:" + id;
    setManualStudents((prev) => prev.filter((m) => m.id !== id));
    removeTransfer(key);
    resetCourseChange(key);
    setExtraCourses((prev) =>
      prev.map((x) => ({
        ...x,
        studentKeys: x.studentKeys.filter((k) => k !== key),
      })),
    );
  };
  // ---- ⑥ 공동교육과정 · 비고 ----
  const addExtraCourse = () =>
    setExtraCourses((prev) => [
      {
        id: newId("ex"),
        kind: "공동교육과정",
        name: "",
        group: "",
        credit: "",
        semester: "",
        memo: "",
        studentKeys: [],
        studentNames: {},
      },
      ...prev,
    ]);
  const updateExtraCourse = (id, patch) =>
    setExtraCourses((prev) =>
      prev.map((x) =>
        x.id === id
          ? {
              ...x,
              ...patch,
            }
          : x,
      ),
    );
  const removeExtraCourse = (id) => setExtraCourses((prev) => prev.filter((x) => x.id !== id));
  const addStudentToExtra = (id, result) =>
    setExtraCourses((prev) =>
      prev.map((x) =>
        x.id === id && !x.studentKeys.includes(result.key)
          ? {
              ...x,
              studentKeys: [...x.studentKeys, result.key],
              studentNames: {
                ...x.studentNames,
                [result.key]: result.name,
              },
            }
          : x,
      ),
    );
  const removeStudentFromExtra = (id, key) =>
    setExtraCourses((prev) =>
      prev.map((x) =>
        x.id === id
          ? {
              ...x,
              studentKeys: x.studentKeys.filter((k) => k !== key),
            }
          : x,
      ),
    );
  // ---- inputs whose student can no longer be found (files re-uploaded, merges changed) ----
  const orphanItems = useMemo(() => {
    if (!curriculumFile) return [];
    const items = [];
    Object.entries(courseChanges).forEach(([k, v]) => {
      if (!resultByKey.has(k))
        items.push({
          id: `change|${k}`,
          kind: "change",
          key: k,
          name: v.name || nameFromStudentKey(k),
          label: "④ 선택과목 변경 내역",
        });
    });
    Object.entries(transferInfo).forEach(([k, v]) => {
      if (!resultByKey.has(k))
        items.push({
          id: `transfer|${k}`,
          kind: "transfer",
          key: k,
          name: v.name || nameFromStudentKey(k),
          label: "⑤ 전입생 정보",
        });
    });
    extraCourses.forEach((x) =>
      x.studentKeys.forEach((k) => {
        if (!resultByKey.has(k))
          items.push({
            id: `extra|${x.id}|${k}`,
            kind: "extra",
            key: k,
            extraId: x.id,
            name: (x.studentNames || {})[k] || nameFromStudentKey(k),
            label: `⑥ ${x.name || "(과목명 없음)"}`,
          });
      }),
    );
    Object.entries(archivedStudents).forEach(([k, v]) => {
      if (!resultByKey.has(k))
        items.push({
          id: `archive|${k}`,
          kind: "archive",
          key: k,
          name: v.name || nameFromStudentKey(k),
          label: "보관",
        });
    });
    return items;
  }, [curriculumFile, resultByKey, courseChanges, transferInfo, extraCourses, archivedStudents]);
  const relinkOrphan = (item, newKey) => {
    const target = resultByKey.get(newKey);
    if (!target || newKey === item.key) return;
    if (item.kind === "change") {
      setCourseChanges((prev) => {
        const copy = {
          ...prev,
        };
        const old = copy[item.key];
        delete copy[item.key];
        if (!old) return copy;
        const ex = copy[newKey];
        copy[newKey] = ex
          ? {
              ...ex,
              name: target.name,
              removed: Array.from(new Set([...ex.removed, ...old.removed])),
              added: [
                ...ex.added,
                ...old.added.filter(
                  (a) => !ex.added.some((b) => entryKey(b.name, b.semester) === entryKey(a.name, a.semester)),
                ),
              ],
              memo: [ex.memo, old.memo].filter(Boolean).join("\n"),
            }
          : {
              ...old,
              name: target.name,
            };
        return copy;
      });
    } else if (item.kind === "transfer") {
      setTransferInfo((prev) => {
        const copy = {
          ...prev,
        };
        const old = copy[item.key];
        delete copy[item.key];
        if (!old) return copy;
        copy[newKey] = copy[newKey]
          ? {
              ...copy[newKey],
              priorCourses: [...copy[newKey].priorCourses, ...old.priorCourses],
            }
          : {
              ...old,
              name: target.name,
            };
        return copy;
      });
    } else if (item.kind === "extra") {
      setExtraCourses((prev) =>
        prev.map((x) =>
          x.id !== item.extraId
            ? x
            : {
                ...x,
                studentKeys: Array.from(new Set(x.studentKeys.map((k) => (k === item.key ? newKey : k)))),
                studentNames: {
                  ...x.studentNames,
                  [newKey]: target.name,
                },
              },
        ),
      );
    } else if (item.kind === "archive") {
      setArchivedStudents((prev) => {
        const copy = {
          ...prev,
        };
        const old = copy[item.key];
        delete copy[item.key];
        if (old)
          copy[newKey] = {
            ...old,
            name: target.name,
          };
        return copy;
      });
    }
  };
  const discardOrphan = (item) => {
    if (item.kind === "change") resetCourseChange(item.key);
    else if (item.kind === "transfer") removeTransfer(item.key);
    else if (item.kind === "extra") removeStudentFromExtra(item.extraId, item.key);
    else if (item.kind === "archive") restoreStudent(item.key);
  };
  // ---- navigation between the student pages ----
  const openChange = (key) => {
    setSelected(null);
    setChangeKey(key);
    setPage("change");
  };
  const openTransfer = (key) => {
    setSelected(null);
    const r = resultByKey.get(key);
    if (r && !transferInfo[key]) registerTransfer(r);
    else setTransferKey(key);
    setPage("transfer");
  };
  const openExtra = () => {
    setSelected(null);
    setPage("extra");
  };
  const openCheckDetail = (key) => {
    setSelected(key);
    setPage("check");
  };
  const handleExport = () => {
    const semList = (list) => list.map((e) => `${e.name}(${formatSemester(e.semester)})`).join(" / ");
    const headers = [
      "학번",
      "학급",
      "이름",
      ...groups,
      "기타 교과(군)",
      "총학점",
      "판정",
      "확인 사유",
      "미이수영역",
      "신청 누락 학기",
      "학기중복선택",
      "학기불일치(편제표와 다른 학기)",
      "인식 못한 과목",
      "전입(본교 첫 학기)",
      "전적교",
      "전적교 이수학점",
      "공동교육과정 등",
      "상담 변경 내역",
      "상담 메모",
      "수강신청 시스템 반영",
    ];
    const rows = activeResults.map((r) => {
      const row = {
        학번: r.studentId,
        학급: r.classLabel,
        이름: r.name,
      };
      groups.forEach((g) => {
        var _a;
        return (row[g] = `${r.groupCredit[g] || 0}/${(_a = reqOverride[g]) !== null && _a !== void 0 ? _a : 0}`);
      });
      row["기타 교과(군)"] = r.otherGroups.map((o) => `${o.group} ${o.credit}`).join(" / ");
      row["총학점"] = r.total;
      row["판정"] = r.pass ? "충족" : "확인필요";
      row["확인 사유"] = r.reasons.map((x) => x.text).join(" / ");
      row["미이수영역"] = r.shortGroups.join(" / ");
      row["신청 누락 학기"] = r.missingSemesters.map(formatSemester).join(" / ");
      row["학기중복선택"] = r.duplicateSelections
        .map((d) => `${d.name}(${d.semesters.map(formatSemester).join(",")})`)
        .join(" / ");
      row["학기불일치(편제표와 다른 학기)"] = r.semesterMismatch
        .map((d) => `${d.name}(${formatSemester(d.semester)} 신청, 편제표 ${d.offered.map(formatSemester).join("·")})`)
        .join(" / ");
      row["인식 못한 과목"] = r.unmatched.join(" / ");
      row["전입(본교 첫 학기)"] = r.isTransfer ? formatSemester(r.firstSem) : "";
      row["전적교"] = r.prevSchool;
      row["전적교 이수학점"] = r.isTransfer ? r.priorCourses.reduce((a, p) => a + p.credit, 0) : "";
      row["공동교육과정 등"] = r.extraEntries.map((x) => `${x.name}(${x.kind}, ${x.credit}학점)`).join(" / ");
      row["상담 변경 내역"] = [
        r.change.added.length ? `추가: ${semList(r.change.added)}` : "",
        r.change.removed.length ? `취소: ${semList(r.change.removed)}` : "",
      ]
        .filter(Boolean)
        .join(" | ");
      row["상담 메모"] = r.change.memo;
      row["수강신청 시스템 반영"] =
        r.change.applyState === "applied"
          ? `반영함(${formatDateYmd(new Date(r.change.appliedAt))})`
          : r.change.applyState === "pending"
            ? "미반영"
            : r.change.applyState === "stale"
              ? "반영 후 다시 바뀜(재반영 필요)"
              : "";
      return row;
    });
    downloadBlob(
      new Blob([toCsv(rows, headers)], {
        type: "text/csv;charset=utf-8",
      }),
      "졸업요건_점검결과.csv",
    );
  };
  const counseledCount = useMemo(() => activeResults.filter((r) => r.hasChanges).length, [activeResults]);
  const tabs = [
    {
      id: "curriculum",
      label: "① 편제표 업로드 (엑셀)",
      ready: true,
      done: !!curriculumFile && curriculumFile.origin !== "builder",
    },
    {
      id: "builder",
      label: "① 편제표 직접 작성 (수작업)",
      ready: true,
      orBefore: true,
      done: (curriculumFile === null || curriculumFile === void 0 ? void 0 : curriculumFile.origin) === "builder",
    },
    {
      id: "roster",
      label: "② 수강신청 · 학생 확인",
      ready: true,
    },
    {
      id: "check",
      label: "③ 졸업요건 점검",
      ready: !!curriculumFile,
    },
    {
      id: "change",
      label: "④ 선택과목 변경",
      ready: !!curriculumFile,
      count: counseledCount,
    },
    {
      id: "transfer",
      label: "⑤ 전입생",
      ready: !!curriculumFile,
      count: Object.keys(transferInfo).length,
    },
    {
      id: "extra",
      label: "⑥ 공동교육과정 · 비고",
      ready: !!curriculumFile,
      count: extraCourses.length,
    },
  ];
  // ---- save/load handlers ----
  const hasWork = !!curriculumFile || regFiles.length > 0 || surveyFiles.length > 0;
  useEffect(() => {
    if (!hasWork) return undefined;
    // 작업 내용은 이 화면(브라우저 메모리)에만 있으므로, 실수로 창을 닫거나 새로고침하기 전에 한 번 묻습니다.
    const handler = (e) => {
      if (keeperRef.current && !keepDirtyRef.current) return undefined; // 브라우저에 최신 내용이 보관되어 있음
      e.preventDefault();
      e.returnValue = "";
      return "";
    };
    window.addEventListener("beforeunload", handler);
    return () => window.removeEventListener("beforeunload", handler);
  }, [hasWork]);
  const buildSnapshot = () => {
    var _a;
    return {
      version: SNAPSHOT_VERSION,
      savedAt: new Date().toISOString(),
      curriculumFile: curriculumFile
        ? {
            name: curriculumFile.name,
            origin: curriculumFile.origin,
            courses: curriculumFile.courses,
            groups: curriculumFile.groups,
            subjectMasterEntries: Array.from(
              ((_a = curriculumFile.subjectMaster) === null || _a === void 0 ? void 0 : _a.entries()) || [],
            ),
          }
        : null,
      reqOverride,
      minCredit,
      regFiles,
      semesterSourceChoice,
      mergeGroups,
      archivedStudents,
      classTeachers,
      builderData,
      priorityGroups,
      admissionYear,
      creativeActivityCredit,
      classMappingFiles,
      courseChanges,
      transferInfo,
      extraCourses,
      manualStudents,
      rosterMode,
      surveyConfig,
      surveyFiles,
      surveyChoice,
      surveyExcluded,
      grade1Electives,
    };
  };
  const keepSummary = () => ({
    curriculum: curriculumFile ? curriculumFile.name : "",
    students: results.length,
    files: regFiles.length + surveyFiles.length,
  });
  useEffect(() => {
    if (!keeperRef.current || !hasWork) return undefined;
    keepDirtyRef.current = true;
    clearTimeout(keepTimerRef.current);
    keepTimerRef.current = setTimeout(async () => {
      const keeper = keeperRef.current;
      if (!keeper) return;
      try {
        const payload = await keeper.encrypt(JSON.stringify(buildSnapshot()));
        await keepWrite({ savedAt: new Date().toISOString(), summary: keepSummary(), payload });
        keepDirtyRef.current = false;
        setKeepState({ savedAt: new Date(), error: "" });
      } catch (e) {
        setKeepState((k) => ({ ...k, error: (e && e.message) || "브라우저에 보관하지 못했습니다." }));
      }
    }, 3000);
    return () => clearTimeout(keepTimerRef.current);
  }, [
    keepPassword,
    curriculumFile,
    reqOverride,
    minCredit,
    regFiles,
    semesterSourceChoice,
    mergeGroups,
    archivedStudents,
    classTeachers,
    builderData,
    priorityGroups,
    admissionYear,
    creativeActivityCredit,
    classMappingFiles,
    courseChanges,
    transferInfo,
    extraCourses,
    manualStudents,
    rosterMode,
    surveyConfig,
    surveyFiles,
    surveyChoice,
    surveyExcluded,
    grade1Electives,
  ]);
  const confirmResume = async () => {
    if (!resumeInfo || !resumePassword) return;
    setResumeBusy(true);
    setResumeError("");
    try {
      const json = await decryptText(resumeInfo.payload, resumePassword);
      restoreSnapshot(JSON.parse(json));
      setLastSavedAt(null); // 브라우저 보관분을 연 것이지 파일로 저장한 것은 아님
      await startKeeping(resumePassword);
      setKeepState({ savedAt: new Date(resumeInfo.savedAt), error: "" });
      setResumeInfo(null);
      setResumePassword("");
    } catch (e) {
      setResumeError(
        e && e.name === "OperationError"
          ? "비밀번호가 올바르지 않습니다."
          : (e && e.message) || "지난 작업을 불러오지 못했습니다.",
      );
    } finally {
      setResumeBusy(false);
    }
  };
  const discardResume = async () => {
    await keepClear();
    setResumeInfo(null);
    setResumePassword("");
    setResumeError("");
  };
  const confirmSave = async () => {
    setSaveError(null);
    try {
      const json = JSON.stringify(buildSnapshot());
      let outText = json;
      const stamp = new Date().toISOString().slice(0, 10);
      let filename = `졸업요건_진행상황_${stamp}.json`;
      if (usePassword) {
        if (!savePassword) return;
        if (!cryptoAvailable)
          throw new Error(
            "이 환경에서는 암호화 저장을 쓸 수 없습니다. 비밀번호 없이 저장하거나, https 주소나 내 컴퓨터의 파일로 열어주세요.",
          );
        outText = JSON.stringify(await encryptText(json, savePassword));
        filename = `졸업요건_진행상황(암호화)_${stamp}.json`;
      }
      downloadBlob(
        new Blob([outText], {
          type: "application/json",
        }),
        filename,
      );
      setLastSavedAt(new Date());
      setSaveModal(false);
      setUsePassword(false);
      setSavePassword("");
    } catch (e) {
      setSaveError(e.message || "저장하지 못했습니다.");
    }
  };
  const restoreSnapshot = (raw) => {
    const s = normalizeSnapshot(raw); // throws a readable error when the file isn't a valid 진행 상황 file
    setCurriculumFile(s.curriculumFile);
    setCurriculumError(null);
    setReqOverride(s.reqOverride);
    setMinCredit(s.minCredit);
    setRegFiles(s.regFiles);
    setSemesterSourceChoice(s.semesterSourceChoice);
    setMergeGroups(s.mergeGroups);
    setArchivedStudents(s.archivedStudents);
    setClassTeachers(s.classTeachers);
    setBuilderData(s.builderData);
    setPriorityGroups(s.priorityGroups);
    setAdmissionYear(s.admissionYear);
    setCreativeActivityCredit(s.creativeActivityCredit);
    setClassMappingFiles(s.classMappingFiles);
    setCourseChanges(s.courseChanges);
    setTransferInfo(s.transferInfo);
    setExtraCourses(s.extraCourses);
    setManualStudents(s.manualStudents);
    setRosterMode(s.rosterMode);
    setSurveyConfig(s.surveyConfig);
    setSurveyFiles(s.surveyFiles);
    setSurveyChoice(s.surveyChoice);
    setSurveyExcluded(s.surveyExcluded);
    setGrade1Electives(s.grade1Electives);
    setSurveyPreview(false);
    setPrintQr(null);
    setPrintList(null);
    // transient UI state from the previous session must not point at students that no longer exist
    setChecked({});
    setSelected(null);
    setSelectedClass(null);
    setSearch("");
    setStatusFilter("all");
    setPrintKeys(null);
    setChangeKey(null);
    setTransferKey(null);
    setArchiveModalKey(null);
    const savedAt = s.savedAt ? new Date(s.savedAt) : null;
    setLastSavedAt(savedAt && !isNaN(savedAt.getTime()) ? savedAt : null); // the loaded file itself is the latest save
    setPage(s.curriculumFile ? "check" : "curriculum");
  };
  const handleLoadFile = async (file) => {
    setLoadError(null);
    let parsed;
    try {
      parsed = JSON.parse(await file.text());
    } catch (e) {
      setLoadError("파일을 읽을 수 없습니다. 이 프로그램의 ‘진행 상황 저장’으로 만든 .json 파일인지 확인해주세요.");
      return;
    }
    if (parsed && parsed.encrypted) {
      if (!cryptoAvailable) {
        setLoadError(
          "이 환경에서는 암호화된 파일을 열 수 없습니다. https 주소나 내 컴퓨터의 파일로 프로그램을 열어주세요.",
        );
        return;
      }
      setLoadPassword("");
      setPendingLoad(parsed);
      return;
    }
    try {
      restoreSnapshot(parsed);
    } catch (e) {
      setLoadError(e.message || "진행 상황을 불러오지 못했습니다.");
    }
  };
  const confirmLoadPassword = async () => {
    let json;
    try {
      json = await decryptText(pendingLoad, loadPassword);
    } catch (e) {
      setLoadError("비밀번호가 올바르지 않거나 파일이 손상되었습니다.");
      return;
    }
    setPendingLoad(null);
    setLoadPassword("");
    try {
      restoreSnapshot(JSON.parse(json));
      setLoadError(null);
    } catch (e) {
      setLoadError(e.message || "진행 상황을 불러오지 못했습니다.");
    }
  };
  const savedLabel = lastSavedAt
    ? `마지막 저장 ${lastSavedAt.toDateString() === new Date().toDateString() ? "" : `${lastSavedAt.getMonth() + 1}/${lastSavedAt.getDate()} `}${String(lastSavedAt.getHours()).padStart(2, "0")}:${String(lastSavedAt.getMinutes()).padStart(2, "0")}`
    : hasWork
      ? "아직 저장하지 않음"
      : "";
  return (
    <div
      className="app-root"
      style={{
        fontFamily: FONT,
        background: PAPER,
        minHeight: "100%",
        color: INK,
      }}
    >
      <div
        style={{
          borderBottom: `1px solid ${LINE}`,
          background: "#fff",
          position: "sticky",
          top: 0,
          zIndex: 10,
        }}
      >
        <div
          style={{
            maxWidth: 1180,
            margin: "0 auto",
            padding: "14px 24px 0",
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              marginBottom: 10,
              gap: 10,
              flexWrap: "wrap",
            }}
          >
            <div
              style={{
                fontSize: 15,
                fontWeight: 700,
              }}
            >
              졸업이수요건 점검
            </div>
            <div
              style={{
                display: "flex",
                gap: 6,
                alignItems: "center",
              }}
            >
              {savedLabel && (
                <span
                  style={{
                    fontSize: 11.5,
                    color: lastSavedAt || keepPassword ? MUTED : WARN,
                  }}
                >
                  {keepPassword
                    ? `브라우저 보관 ${keepState.error ? "실패" : keepState.savedAt ? `${String(keepState.savedAt.getHours()).padStart(2, "0")}:${String(keepState.savedAt.getMinutes()).padStart(2, "0")}` : "중…"} · ${lastSavedAt ? savedLabel : "파일 저장 안 함"}`
                    : savedLabel}
                </span>
              )}
              <button
                onClick={() => {
                  setSaveError(null);
                  if (keepPassword) {
                    setUsePassword(true);
                    setSavePassword(keepPassword);
                  }
                  setSaveModal(true);
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 5,
                  fontSize: 12.5,
                  fontWeight: 600,
                  color: ACCENT,
                  background: ACCENT_BG,
                  border: "none",
                  borderRadius: 7,
                  padding: "6px 11px",
                  cursor: "pointer",
                }}
              >
                <Save size={13} />
                {" \uC9C4\uD589 \uC0C1\uD669 \uC800\uC7A5"}
              </button>
              <button
                onClick={() => {
                  var _a;
                  if (
                    hasWork &&
                    !window.confirm(
                      "불러오면 지금 화면의 작업 내용이 파일의 내용으로 바뀝니다. 저장하지 않은 내용은 사라집니다. 계속할까요?",
                    )
                  )
                    return;
                  (_a = loadInputRef.current) === null || _a === void 0 ? void 0 : _a.click();
                }}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 5,
                  fontSize: 12.5,
                  fontWeight: 600,
                  color: INK,
                  background: PAPER,
                  border: `1px solid ${LINE}`,
                  borderRadius: 7,
                  padding: "6px 11px",
                  cursor: "pointer",
                }}
              >
                <FileUp size={13} />
                {" \uBD88\uB7EC\uC624\uAE30"}
              </button>
              <input
                ref={loadInputRef}
                type="file"
                accept=".json,application/json"
                style={{
                  display: "none",
                }}
                onChange={(e) => {
                  var _a;
                  const f = (_a = e.target.files) === null || _a === void 0 ? void 0 : _a[0];
                  if (f) handleLoadFile(f);
                  e.target.value = "";
                }}
              />
            </div>
          </div>
          <div
            style={{
              display: "flex",
              gap: 2,
              flexWrap: "wrap",
            }}
          >
            {tabs.map((t) => (
              <React.Fragment key={t.id}>
                {t.orBefore && (
                  <span
                    style={{
                      alignSelf: "center",
                      fontSize: 11.5,
                      color: "#B8B4A9",
                      padding: "0 2px",
                    }}
                  >
                    또는
                  </span>
                )}
                <button
                  onClick={() => t.ready && setPage(t.id)}
                  title={t.ready ? "" : "편제표를 먼저 올려주세요"}
                  style={{
                    border: "none",
                    background: "none",
                    cursor: t.ready ? "pointer" : "default",
                    padding: "8px 10px",
                    fontSize: 13,
                    fontWeight: 600,
                    color: page === t.id ? ACCENT : t.ready ? MUTED : "#B8B4A9",
                    borderBottom: page === t.id ? `2px solid ${ACCENT}` : "2px solid transparent",
                    fontFamily: FONT,
                  }}
                >
                  {t.label}
                  {t.done && (
                    <span
                      style={{
                        color: OK,
                        marginLeft: 5,
                      }}
                    >
                      ✓
                    </span>
                  )}
                  {t.count > 0 && (
                    <span
                      style={{
                        marginLeft: 5,
                        fontSize: 11,
                        fontWeight: 700,
                        color: ACCENT,
                        background: ACCENT_BG,
                        borderRadius: 999,
                        padding: "1px 6px",
                      }}
                    >
                      {t.count}
                    </span>
                  )}
                </button>
              </React.Fragment>
            ))}
          </div>
        </div>
      </div>
      <div
        style={{
          maxWidth: 1180,
          margin: "0 auto",
          padding: "24px 24px 60px",
        }}
      >
        {page === "curriculum" && (
          <CurriculumPage
            curriculumFile={curriculumFile}
            curriculumError={curriculumError}
            handleCurriculum={handleCurriculum}
            clearCurriculum={() => {
              setCurriculumFile(null);
              setCurriculumError(null);
            }}
            minCredit={minCredit}
            setMinCredit={setMinCredit}
            reqOverride={reqOverride}
            setReqOverride={setReqOverride}
            showReq={showReq}
            setShowReq={setShowReq}
            groups={groups}
            onNext={() => setPage("roster")}
            onGoBuilder={() => setPage("builder")}
            priorityGroups={priorityGroups}
            setPriorityGroups={setPriorityGroups}
            admissionYear={admissionYear}
            setAdmissionYear={setAdmissionYear}
            creativeActivityCredit={creativeActivityCredit}
            setCreativeActivityCredit={setCreativeActivityCredit}
            curriculumIssues={curriculumIssues}
          />
        )}
        {page === "builder" && (
          <BuilderPage builderData={builderData} setBuilderData={setBuilderData} onApply={applyBuilderData} />
        )}
        {page === "roster" && (
          <RosterPage
            regFiles={regFiles}
            regFileInfos={regFileInfos}
            handleRegFiles={handleRegFiles}
            removeRegFile={removeRegFile}
            relabelRegFile={relabelRegFile}
            setRegFileSemester={setRegFileSemester}
            moveRegFile={moveRegFile}
            semesterConflicts={semesterConflicts}
            setSemesterSource={setSemesterSource}
            rawRecords={rawRecords}
            ambiguousNames={ambiguousNames}
            identityKeyFor={identityKeyFor}
            mergeGroups={mergeGroups}
            checked={checked}
            toggleCheck={toggleCheck}
            mergeChecked={mergeChecked}
            unmerge={unmerge}
            unresolvedCount={unresolvedCount}
            studentCount={studentGroups.length}
            onNext={() => setPage("check")}
            canProceed={!!curriculumFile}
            semesterSummary={semesterSummary}
            hasCurriculum={!!curriculumFile}
            classMappingFiles={classMappingFiles}
            handleClassMappingFiles={handleClassMappingFiles}
            removeClassMappingFile={removeClassMappingFile}
            mappedResolvedCount={mappedResolvedCount}
            hakjeomLinkedCount={hakjeomLinkedCount}
            rosterMode={rosterMode}
            setRosterMode={setRosterMode}
            surveyPanel={
              <SurveyPanel
                config={surveyConfig}
                setConfig={setSurveyConfig}
                poolDefs={semesterPoolDefs}
                payload={surveyPayload}
                problems={surveyIssues}
                hasCurriculum={!!curriculumFile}
                admissionYear={admissionYear}
                onPreview={() => setSurveyPreview(true)}
                onPrintQr={printSurveyQr}
                surveyFiles={surveyFiles}
                onSurveyFiles={handleSurveyFiles}
                onFetchLive={fetchSurveyLive}
                surveyFetch={surveyFetch}
                removeSurveyFile={removeSurveyFile}
                surveyData={surveyData}
                surveyChoice={surveyChoice}
                setSurveyChoice={setSurveyChoiceFor}
                surveyExcluded={surveyExcluded}
                toggleSurveyExcluded={toggleSurveyExcluded}
                onGoTransfer={() => setPage("transfer")}
                toggleSurveyFileDisabled={toggleSurveyFileDisabled}
              />
            }
            surveyFileCount={surveyFiles.length}
            sourcePanel={
              <SemesterSourcePanel
                semesterSummary={semesterSummary}
                providers={semesterProviders}
                chosenSourceBySem={chosenSourceBySem}
                explicitChoice={semesterSourceChoice}
                regFiles={regFiles}
                regFileInfos={regFileInfos}
                surveyData={surveyData}
                setSemesterSource={setSemesterSource}
                grade1={grade1Status}
                setGrade1Answer={setGrade1Electives}
                onGoCurriculum={() =>
                  setPage(curriculumFile && curriculumFile.origin === "builder" ? "builder" : "curriculum")
                }
              />
            }
            sourceSemesters={sourceSemesters}
            onToggleRegFileDisabled={toggleRegFileDisabled}
          />
        )}
        {page === "check" && (
          <CheckPage
            curriculumFile={curriculumFile}
            groups={orderedGroups}
            priorityGroups={priorityGroups}
            reqOverride={reqOverride}
            minCredit={minCredit}
            results={activeResults}
            allResults={results}
            filtered={filtered}
            passCount={passCount}
            failCount={failCount}
            search={search}
            setSearch={setSearch}
            statusFilter={statusFilter}
            setStatusFilter={setStatusFilter}
            statusFilters={checkFilters}
            selected={selected}
            setSelected={setSelected}
            selectedResult={selectedResult}
            handleExport={handleExport}
            classSummary={classSummary}
            selectedClass={selectedClass}
            setSelectedClass={setSelectedClass}
            archivedResults={archivedResults}
            archiveModalKey={archiveModalKey}
            setArchiveModalKey={setArchiveModalKey}
            archiveReason={archiveReason}
            setArchiveReason={setArchiveReason}
            confirmArchive={confirmArchive}
            restoreStudent={restoreStudent}
            printOne={printOne}
            printClass={printClass}
            classTeachers={classTeachers}
            setClassTeachers={setClassTeachers}
            onOpenChange={openChange}
            onOpenTransfer={openTransfer}
            onOpenExtra={openExtra}
            orphanItems={orphanItems}
            relinkOrphan={relinkOrphan}
            discardOrphan={discardOrphan}
            surveyMissingByClass={surveyMissingByClass}
            sourceIssues={{
              unresolvedSems: unresolvedSourceSems,
              grade1Warn: grade1Status.level === "warn",
            }}
            onGoRoster={() => setPage("roster")}
            setChangeApplied={setChangeApplied}
            onPrintList={printProblemList}
          />
        )}
        {page === "change" && (
          <ChangePage
            results={activeResults}
            groups={orderedGroups}
            reqOverride={reqOverride}
            priorityGroups={priorityGroups}
            poolDefs={semesterPoolDefs}
            selectedKey={changeKey}
            setSelectedKey={setChangeKey}
            toggleCourseChange={toggleCourseChange}
            setChangeMemo={setChangeMemo}
            resetCourseChange={resetCourseChange}
            setChangeApplied={setChangeApplied}
            onOpenCheck={openCheckDetail}
            onOpenTransfer={openTransfer}
            printOne={printOne}
          />
        )}
        {page === "transfer" && (
          <TransferPage
            results={activeResults}
            transferInfo={transferInfo}
            courses={(curriculumFile === null || curriculumFile === void 0 ? void 0 : curriculumFile.courses) || []}
            groups={orderedGroups}
            reqOverride={reqOverride}
            priorityGroups={priorityGroups}
            selectedKey={transferKey}
            setSelectedKey={setTransferKey}
            registerTransfer={registerTransfer}
            updateTransfer={updateTransfer}
            removeTransfer={removeTransfer}
            manualStudents={manualStudents}
            addManualStudent={addManualStudent}
            updateManualStudent={updateManualStudent}
            removeManualStudent={removeManualStudent}
            onOpenCheck={openCheckDetail}
            onOpenChange={openChange}
            neisFiles={neisFiles}
            handleNeisFiles={handleNeisFiles}
            removeNeisFile={removeNeisFile}
            applyNeisRows={applyNeisRows}
          />
        )}
        {page === "extra" && (
          <ExtraPage
            results={activeResults}
            allResults={results}
            extraCourses={extraCourses}
            groups={orderedGroups}
            courses={(curriculumFile === null || curriculumFile === void 0 ? void 0 : curriculumFile.courses) || []}
            addExtraCourse={addExtraCourse}
            updateExtraCourse={updateExtraCourse}
            removeExtraCourse={removeExtraCourse}
            addStudentToExtra={addStudentToExtra}
            removeStudentFromExtra={removeStudentFromExtra}
            onOpenCheck={openCheckDetail}
          />
        )}
      </div>
      <footer
        style={{
          borderTop: `1px solid ${LINE}`,
          padding: "18px 24px 28px",
          textAlign: "center",
          fontSize: 11.5,
          color: MUTED,
          lineHeight: 1.8,
        }}
      >
        <div
          style={{
            maxWidth: 920,
            margin: "0 auto",
          }}
        >
          <div>
            {"\uD559\uC0DD \uC774\uB984\u00B7\uD559\uBC88\u00B7\uC120\uD0DD \uACFC\uBAA9 \uB4F1 "}
            <b
              style={{
                color: INK,
              }}
            >
              개인정보
            </b>
            가 담긴 자료입니다.
          </div>
          <div>「개인정보 보호법」에 따라 졸업 이수 점검 업무에만 사용하고, 끝나면 파기해 주세요.</div>
          <div>저장 파일·출력물은 안전하게 보관하고, 무단 유출·배포를 금합니다.</div>
          <div>
            올린 자료는 인터넷으로 보내지 않고 이 컴퓨터에서만 처리합니다. (기초조사 응답은 선생님 구글 시트에 저장)
          </div>
          <div
            style={{
              marginTop: 8,
              color: INK,
              fontWeight: 700,
            }}
          >
            2026 고교학점제 종합지원단 학점이수관리팀 제작
          </div>
          <div>Copyright © 2026. All rights reserved. 무단 복제·수정·재배포 및 상업적 이용을 금합니다.</div>
          <div style={{ marginTop: 4 }}>
            프로그램 업데이트 {__BUILD_INFO__.builtAt}
            {__BUILD_INFO__.commit ? ` (${__BUILD_INFO__.commit})` : ""}
          </div>
        </div>
      </footer>
      {surveyPreview && <SurveyPreviewModal html={surveyPreviewHtml} onClose={() => setSurveyPreview(false)} />}
      {cryptoAvailable && hasWork && !keepPassword && !keepDeclined && !resumeInfo && (
        <KeepOfferBar
          value={keepDraft}
          onChange={setKeepDraft}
          onKeep={() => keepDraft.length >= 4 && startKeeping(keepDraft)}
          onDecline={declineKeeping}
        />
      )}
      {resumeInfo && (
        <ModalShell onClose={() => {}}>
          <div
            style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 15, fontWeight: 700, marginBottom: 6 }}
          >
            <Lock size={16} />
            {" 이 브라우저에 지난 작업이 있습니다"}
          </div>
          <div style={{ fontSize: 12.5, color: MUTED, marginBottom: 12, lineHeight: 1.6 }}>
            {(() => {
              const d = new Date(resumeInfo.savedAt);
              const sm = resumeInfo.summary || {};
              return `${d.getMonth() + 1}월 ${d.getDate()}일 ${String(d.getHours()).padStart(2, "0")}:${String(d.getMinutes()).padStart(2, "0")} 보관${sm.curriculum ? ` · 편제표 ${sm.curriculum}` : ""}${sm.students ? ` · 학생 ${sm.students}명` : ""}`;
            })()}
            <br />
            보관할 때 정한 비밀번호를 넣으면 이어서 합니다. 비밀번호를 잊었다면 되찾을 수 없으니, 파일로 저장한 것이
            있으면 그 파일을 불러오세요.
          </div>
          <input
            type="password"
            value={resumePassword}
            onChange={(e) => setResumePassword(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter") confirmResume();
            }}
            placeholder="비밀번호 입력"
            autoFocus={true}
            style={{
              width: "100%",
              boxSizing: "border-box",
              padding: "9px 11px",
              borderRadius: 7,
              border: `1px solid ${LINE}`,
              fontSize: 13,
              marginBottom: 10,
            }}
          />
          {resumeError && <div style={{ fontSize: 12.5, color: WARN, marginBottom: 10 }}>{resumeError}</div>}
          <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
            <button
              onClick={() => {
                if (window.confirm("브라우저에 보관된 지난 작업을 지웁니다. 되돌릴 수 없습니다. 계속할까요?"))
                  discardResume();
              }}
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: WARN,
                background: "none",
                border: "none",
                padding: "8px 4px",
                cursor: "pointer",
              }}
            >
              지우고 새로 시작
            </button>
            <button
              disabled={!resumePassword || resumeBusy}
              onClick={confirmResume}
              style={{
                fontSize: 13,
                fontWeight: 700,
                color: "#fff",
                background: !resumePassword || resumeBusy ? "#B8B4A9" : ACCENT,
                border: "none",
                borderRadius: 7,
                padding: "8px 16px",
                cursor: !resumePassword || resumeBusy ? "default" : "pointer",
              }}
            >
              {resumeBusy ? "여는 중…" : "이어하기"}
            </button>
          </div>
        </ModalShell>
      )}
      {saveModal && (
        <ModalShell onClose={() => setSaveModal(false)}>
          <div
            style={{
              fontSize: 15,
              fontWeight: 700,
              marginBottom: 6,
            }}
          >
            진행 상황 저장
          </div>
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              marginBottom: 16,
              lineHeight: 1.55,
            }}
          >
            편제표, 수강신청 파일 내용, 학생 기초조사 설정·응답, 동명이인 병합, 보관, ④ 선택과목 변경, ⑤ 전입생, ⑥
            공동교육과정 입력까지 모두 한 파일(.json)로 내려받습니다. 서버에는 아무것도 저장되지 않으며, 다음에
            ‘불러오기’로 이어서 작업할 수 있습니다. 파일 저장이 정식 저장이고, 브라우저 임시 보관은 이 컴퓨터·이
            브라우저에서만 남습니다.
          </div>
          {cryptoAvailable && (
            <div
              style={{
                fontSize: 12.3,
                color: MUTED,
                background: PAPER,
                borderRadius: 8,
                padding: "8px 10px",
                marginBottom: 14,
                display: "flex",
                gap: 8,
                alignItems: "center",
                flexWrap: "wrap",
              }}
            >
              <span style={{ flex: 1, minWidth: 200 }}>
                {keepPassword
                  ? `이 브라우저 임시 보관: 켜짐 (비밀번호로 잠금 · ${KEEP_MAX_AGE_DAYS}일 뒤 자동 삭제)`
                  : "이 브라우저 임시 보관: 꺼짐"}
              </span>
              {keepPassword ? (
                <button
                  onClick={() => {
                    if (
                      window.confirm(
                        "이 브라우저에 보관된 자료를 지우고 자동 보관을 끕니다. 화면의 작업 내용은 그대로 남습니다. 계속할까요?",
                      )
                    )
                      stopKeeping();
                  }}
                  style={{ ...buttonStyle("ghost"), padding: "5px 10px" }}
                >
                  끄고 지우기
                </button>
              ) : (
                <button
                  onClick={() => {
                    setKeepDeclined(false);
                    try {
                      sessionStorage.removeItem("keepDeclined");
                    } catch (e) {
                      // 무시
                    }
                    setSaveModal(false);
                  }}
                  style={{ ...buttonStyle("ghost"), padding: "5px 10px" }}
                >
                  켜기
                </button>
              )}
            </div>
          )}
          <label
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              fontSize: 13,
              marginBottom: 10,
              cursor: cryptoAvailable ? "pointer" : "default",
              color: cryptoAvailable ? INK : MUTED,
            }}
          >
            <input
              type="checkbox"
              disabled={!cryptoAvailable}
              checked={usePassword}
              onChange={(e) => setUsePassword(e.target.checked)}
            />
            비밀번호로 보호하기 (학생 개인정보가 포함되어 있어 권장합니다)
          </label>
          {!cryptoAvailable && (
            <div
              style={{
                fontSize: 11.5,
                color: WARN,
                marginBottom: 10,
              }}
            >
              이 주소(http)에서는 브라우저가 암호화 기능을 막아 비밀번호 저장을 쓸 수 없습니다.
            </div>
          )}
          {usePassword && (
            <input
              type="password"
              value={savePassword}
              onChange={(e) => setSavePassword(e.target.value)}
              placeholder="비밀번호 입력 (잊어버리면 파일을 열 수 없습니다)"
              style={{
                width: "100%",
                boxSizing: "border-box",
                padding: "9px 11px",
                borderRadius: 7,
                border: `1px solid ${LINE}`,
                fontSize: 13,
                marginBottom: 14,
              }}
            />
          )}
          {saveError && (
            <div
              style={{
                fontSize: 12.5,
                color: WARN,
                marginBottom: 10,
                lineHeight: 1.5,
              }}
            >
              {saveError}
            </div>
          )}
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              gap: 8,
              marginTop: 6,
            }}
          >
            <button
              onClick={() => setSaveModal(false)}
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: MUTED,
                background: "none",
                border: "none",
                padding: "8px 12px",
                cursor: "pointer",
              }}
            >
              취소
            </button>
            <button
              disabled={usePassword && !savePassword}
              onClick={confirmSave}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                fontSize: 13,
                fontWeight: 700,
                color: "#fff",
                background: usePassword && !savePassword ? "#B8B4A9" : ACCENT,
                border: "none",
                borderRadius: 7,
                padding: "8px 16px",
                cursor: usePassword && !savePassword ? "default" : "pointer",
              }}
            >
              <Save size={14} />
              {" \uD30C\uC77C\uB85C \uC800\uC7A5"}
            </button>
          </div>
        </ModalShell>
      )}
      {pendingLoad && (
        <ModalShell
          onClose={() => {
            setPendingLoad(null);
            setLoadPassword("");
            setLoadError(null);
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              fontSize: 15,
              fontWeight: 700,
              marginBottom: 6,
            }}
          >
            <Lock size={16} />
            {" \uBE44\uBC00\uBC88\uD638\uB85C \uBCF4\uD638\uB41C \uD30C\uC77C"}
          </div>
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              marginBottom: 14,
            }}
          >
            이 파일을 저장할 때 설정한 비밀번호를 입력해주세요.
          </div>
          <input
            type="password"
            value={loadPassword}
            onChange={(e) => setLoadPassword(e.target.value)}
            onKeyDown={(e) => {
              if (e.key === "Enter" && loadPassword) confirmLoadPassword();
            }}
            placeholder="비밀번호 입력"
            autoFocus={true}
            style={{
              width: "100%",
              boxSizing: "border-box",
              padding: "9px 11px",
              borderRadius: 7,
              border: `1px solid ${LINE}`,
              fontSize: 13,
              marginBottom: 10,
            }}
          />
          {loadError && (
            <div
              style={{
                fontSize: 12.5,
                color: WARN,
                marginBottom: 10,
              }}
            >
              {loadError}
            </div>
          )}
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              gap: 8,
            }}
          >
            <button
              onClick={() => {
                setPendingLoad(null);
                setLoadPassword("");
                setLoadError(null);
              }}
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: MUTED,
                background: "none",
                border: "none",
                padding: "8px 12px",
                cursor: "pointer",
              }}
            >
              취소
            </button>
            <button
              disabled={!loadPassword}
              onClick={confirmLoadPassword}
              style={{
                fontSize: 13,
                fontWeight: 700,
                color: "#fff",
                background: !loadPassword ? "#B8B4A9" : ACCENT,
                border: "none",
                borderRadius: 7,
                padding: "8px 16px",
                cursor: !loadPassword ? "default" : "pointer",
              }}
            >
              불러오기
            </button>
          </div>
        </ModalShell>
      )}
      {loadError && !pendingLoad && (
        <ModalShell onClose={() => setLoadError(null)}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 8,
              fontSize: 14,
              fontWeight: 700,
              color: WARN,
              marginBottom: 8,
            }}
          >
            <AlertCircle size={16} />
            {" \uBD88\uB7EC\uC624\uAE30 \uC2E4\uD328"}
          </div>
          <div
            style={{
              fontSize: 13,
              color: INK,
              lineHeight: 1.55,
            }}
          >
            {loadError}
          </div>
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              marginTop: 14,
            }}
          >
            <button onClick={() => setLoadError(null)} style={buttonStyle("primary")}>
              확인
            </button>
          </div>
        </ModalShell>
      )}
      <PrintRoot
        results={printResults}
        groups={orderedGroups}
        priorityGroups={priorityGroups}
        reqOverride={reqOverride}
        curriculumFile={curriculumFile}
        classTeachers={classTeachers}
        qrSheet={printQr}
        listSheet={printList}
      />
      <style>{`
        @media print {
          .app-root > *:not(#print-root) { display: none !important; }
          #print-root { display: block !important; }
        }
        #print-root { display: none; }
        .two-col { display: grid; grid-template-columns: minmax(260px, 320px) minmax(0, 1fr); gap: 16px; align-items: start; }
        @media (max-width: 900px) { .two-col { grid-template-columns: minmax(0, 1fr); } }
      `}</style>
    </div>
  );
}
function GraduationChecker() {
  return (
    <ErrorBoundary>
      <GraduationCheckerInner />
    </ErrorBoundary>
  );
}
// 작업이 생기면 오른쪽 아래에 나오는 띠: 비밀번호를 정하면 이 브라우저에 자동 보관을 시작합니다
function KeepOfferBar({ value, onChange, onKeep, onDecline }) {
  const short = value.length > 0 && value.length < 4;
  return (
    <div
      style={{
        position: "fixed",
        right: 16,
        bottom: 16,
        zIndex: 50,
        width: 340,
        maxWidth: "calc(100vw - 32px)",
        background: "#fff",
        border: `1px solid ${ACCENT}`,
        borderRadius: 10,
        boxShadow: "0 6px 24px rgba(28,35,51,0.18)",
        padding: "12px 14px",
        fontSize: 12.5,
        lineHeight: 1.55,
      }}
    >
      <div style={{ fontWeight: 800, fontSize: 13, marginBottom: 4 }}>이 브라우저에 임시 보관할까요?</div>
      <div style={{ color: MUTED, marginBottom: 8 }}>
        실수로 창을 닫아도 이어서 할 수 있게, 작업 내용을 비밀번호로 잠가 이 브라우저에 자동으로 남겨 둡니다.{" "}
        {KEEP_MAX_AGE_DAYS}일 뒤 자동으로 지워집니다. <b style={{ color: INK }}>공용 컴퓨터라면 [안 함]</b>을 누르세요.
      </div>
      <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
        <input
          type="password"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") onKeep();
          }}
          placeholder="비밀번호 (4자 이상)"
          style={{ ...inputStyle, flex: 1, minWidth: 0, padding: "7px 9px", fontSize: 13 }}
        />
        <button
          disabled={value.length < 4}
          onClick={onKeep}
          style={buttonStyle(value.length < 4 ? "disabled" : "primary", { padding: "7px 12px" })}
        >
          보관하기
        </button>
        <button
          onClick={onDecline}
          style={{ ...buttonStyle("ghost"), border: "none", color: MUTED, padding: "7px 8px" }}
        >
          안 함
        </button>
      </div>
      {short && <div style={{ color: WARN, marginTop: 4 }}>4자 이상으로 정해 주세요.</div>}
    </div>
  );
}
function ModalShell({ onClose, children }) {
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(28,35,51,0.35)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 60,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          width: 380,
          maxWidth: "90vw",
          background: "#fff",
          borderRadius: 12,
          padding: 22,
          fontFamily: FONT,
          boxShadow: "0 12px 32px rgba(0,0,0,0.18)",
        }}
      >
        {children}
      </div>
    </div>
  );
}
// Printable summary for one student. Always mounted (hidden on screen), and shown via a plain
// @media print rule using normal block flow — NOT position:absolute — so multi-page content
// paginates correctly instead of overlapping across page breaks.
// Everything that feeds the credit total is listed here (selected electives, 전입 이전 과목, 공동교육과정 등, 기타 반영
// 과목, 상담 후 변경), so the sums on the paper can always be traced back to a line on the same paper.
function StudentPrintCard({ result, groups, priorityGroups, reqOverride, curriculumFile, classTeachers, pageBreak }) {
  var _a, _b, _c, _d, _e, _f;
  // group the selection detail by grade year (2학년 = 2-1+2-2, 3학년 = 3-1+3-2), one column per semester
  const byGradeYear = new Map();
  (result.pools || []).forEach((p) => {
    const m = /^(\d)-(\d)$/.exec(p.sem || "");
    const grade = m ? m[1] : "?";
    if (!byGradeYear.has(grade))
      byGradeYear.set(grade, {
        bySem: new Map(),
        deficientNotes: [],
      });
    const entry = byGradeYear.get(grade);
    if (!entry.bySem.has(p.sem)) entry.bySem.set(p.sem, []);
    const rows = entry.bySem.get(p.sem);
    p.takenCourses.forEach((c) =>
      rows.push({
        group: c.group,
        name: c.name,
        credit: c.credit,
      }),
    );
    if (p.deficient)
      entry.deficientNotes.push(
        `${formatSemester(p.sem)} ${p.label}: ${p.takenCount}/${p.requiredN} 선택 (${p.requiredN - p.takenCount}개 부족)`,
      );
  });
  const gradeYears = Array.from(byGradeYear.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([grade, v]) => ({
      grade,
      deficientNotes: v.deficientNotes,
      semColumns: Array.from(v.bySem.entries())
        .sort((a, b) => compareSemester(a[0], b[0]))
        .map(([sem, rows]) => ({
          sem,
          rows: [...rows].sort(
            (a, b) => b.credit - a.credit || groupPriorityIndex(a.group) - groupPriorityIndex(b.group),
          ),
        })),
    }));
  const teacherName =
    (classTeachers === null || classTeachers === void 0 ? void 0 : classTeachers[result.classLabel]) || "";
  const blocking = (result.reasons || []).filter((x) => x.blocking);
  const otherTaken = (result.taken || []).filter((t) => !t.selectGroup && !t.schoolDesignated);
  const th = {
    textAlign: "left",
    padding: "3px 4px",
    fontSize: 10.5,
    color: MUTED,
    fontWeight: 700,
  };
  const td = {
    padding: "3px 4px",
  };
  const table = (head, rows) => (
    <table
      style={{
        borderCollapse: "collapse",
        width: "100%",
        fontSize: 11.5,
      }}
    >
      <thead>
        <tr>
          {head.map((h, i) => (
            <th
              key={i}
              style={{
                ...th,
                textAlign: i === head.length - 1 ? "right" : "left",
              }}
            >
              {h}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {rows.map((cells, ri) => (
          <tr
            key={ri}
            style={{
              borderTop: `1px solid ${LINE}`,
            }}
          >
            {cells.map((c, ci) => (
              <td
                key={ci}
                style={{
                  ...td,
                  textAlign: ci === cells.length - 1 ? "right" : "left",
                }}
              >
                {c}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  );
  const section = (title, children) => (
    <div
      style={{
        marginTop: 14,
        border: `1px solid ${LINE}`,
        borderRadius: 8,
        padding: 12,
        breakInside: "avoid",
      }}
    >
      <div
        style={{
          fontSize: 13,
          fontWeight: 700,
          marginBottom: 6,
        }}
      >
        {title}
      </div>
      {children}
    </div>
  );
  return (
    <div
      style={{
        fontFamily: FONT,
        color: INK,
        padding: 24,
        breakAfter: pageBreak ? "page" : "auto",
      }}
    >
      <div
        style={{
          fontSize: 20,
          fontWeight: 700,
          textAlign: "center",
        }}
      >
        교육과정 이수 확인서
      </div>
      <div
        style={{
          fontSize: 13,
          color: MUTED,
          marginTop: 6,
          textAlign: "right",
        }}
      >
        {result.classLabel}
        {" \u00B7 "}
        {result.studentId}
        {" \u00B7 "}
        {result.name}
        {result.isTransfer && ` · 전입(본교 ${formatSemester(result.firstSem)}부터)`}
      </div>
      <div
        style={{
          fontSize: 13,
          fontWeight: 700,
          marginTop: 12,
          color: result.pass ? OK : WARN,
        }}
      >
        {result.pass ? "졸업요건 충족" : "확인 필요"}
        {" \u00B7 \uCD1D "}
        {result.total}학점
      </div>
      {blocking.length > 0 && (
        <div
          style={{
            marginTop: 6,
            fontSize: 11.5,
            color: WARN,
            lineHeight: 1.6,
          }}
        >
          {blocking.map((x, i) => (
            <div key={i}>
              {"\u26A0 "}
              {x.text}
            </div>
          ))}
        </div>
      )}
      {((_a = result.missingSemesters) === null || _a === void 0 ? void 0 : _a.length) > 0 && (
        <div
          style={{
            marginTop: 4,
            fontSize: 11.5,
            color: WARN,
          }}
        >
          {"\u26A0 \uC218\uAC15\uC2E0\uCCAD \uAE30\uB85D\uC774 \uC5C6\uB294 \uD559\uAE30: "}
          {result.missingSemesters.map(formatSemester).join(", ")}
        </div>
      )}
      <div
        style={{
          marginTop: 16,
          fontSize: 13,
          fontWeight: 700,
        }}
      >
        교과영역별 이수학점
      </div>
      <div
        style={{
          marginTop: 6,
          border: `1px solid ${LINE}`,
          borderRadius: 8,
          padding: 12,
        }}
      >
        <table
          style={{
            borderCollapse: "collapse",
            width: "100%",
            fontSize: 11.5,
          }}
        >
          <tbody>
            {groups.map((g, gi) => {
              var _a;
              const cur = result.groupCredit[g] || 0;
              const req = (_a = reqOverride[g]) !== null && _a !== void 0 ? _a : 0;
              const short = cur < req;
              const priority = !!priorityGroups[g];
              return (
                <tr
                  key={g}
                  style={{
                    borderTop: gi === 0 ? "none" : `1px solid ${LINE}`,
                    background: priority ? "#FDF3C7" : "transparent",
                  }}
                >
                  <td
                    style={{
                      padding: "3px 4px",
                      fontWeight: priority ? 700 : 400,
                    }}
                  >
                    {g}
                  </td>
                  <td
                    style={{
                      padding: "3px 4px",
                      textAlign: "right",
                      fontWeight: 700,
                      color: short ? WARN : INK,
                    }}
                  >
                    {cur}
                    {" / "}
                    {req}
                    {short ? ` (${req - cur}학점 부족)` : ""}
                  </td>
                </tr>
              );
            })}
            {(result.otherGroups || []).map((o) => (
              <tr
                key={o.group}
                style={{
                  borderTop: `1px solid ${LINE}`,
                }}
              >
                <td
                  style={{
                    padding: "3px 4px",
                    color: MUTED,
                  }}
                >
                  {o.group}
                  {" (\uD544\uC218\uC774\uC218 \uAE30\uC900 \uC5C6\uC74C)"}
                </td>
                <td
                  style={{
                    padding: "3px 4px",
                    textAlign: "right",
                  }}
                >
                  {o.credit}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {gradeYears.length > 0 && (
        <div
          style={{
            marginTop: 18,
          }}
        >
          <div
            style={{
              fontSize: 14,
              fontWeight: 700,
            }}
          >
            선택과목 이수 현황
          </div>
          {gradeYears.map(({ grade, semColumns, deficientNotes }) => (
            <div
              key={grade}
              style={{
                marginTop: 10,
                border: `1px solid ${LINE}`,
                borderRadius: 8,
                padding: 12,
                breakInside: "avoid",
              }}
            >
              <div
                style={{
                  fontSize: 13,
                  fontWeight: 700,
                  marginBottom: 6,
                }}
              >
                {grade}학년
              </div>
              {deficientNotes.length > 0 && (
                <div
                  style={{
                    marginBottom: 8,
                  }}
                >
                  {deficientNotes.map((n, i) => (
                    <div
                      key={i}
                      style={{
                        fontSize: 11.5,
                        color: WARN,
                        fontWeight: 600,
                      }}
                    >
                      {"\u26A0 "}
                      {n}
                    </div>
                  ))}
                </div>
              )}
              <div
                style={{
                  display: "grid",
                  gridTemplateColumns: `repeat(${Math.max(semColumns.length, 1)}, 1fr)`,
                  gap: 16,
                }}
              >
                {semColumns.map((col) => (
                  <div key={col.sem}>
                    <div
                      style={{
                        fontSize: 12,
                        fontWeight: 700,
                        marginBottom: 4,
                      }}
                    >
                      {formatSemester(col.sem)}
                    </div>
                    {col.rows.length > 0 ? (
                      table(
                        ["교과(군)", "과목명", "학점"],
                        col.rows.map((r) => [r.group, r.name, r.credit]),
                      )
                    ) : (
                      <div
                        style={{
                          fontSize: 11.5,
                          color: MUTED,
                        }}
                      >
                        선택한 과목이 없습니다.
                      </div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      )}
      {result.isTransfer &&
        section(
          `전입 이전 이수 과목${result.prevSchool ? ` (전적교: ${result.prevSchool})` : ""}`,
          result.priorCourses.length > 0 ? (
            table(
              ["학기", "교과(군)", "과목명", "학점"],
              result.priorCourses.map((p) => [formatSemester(p.semester), p.group, p.name, p.credit]),
            )
          ) : (
            <div
              style={{
                fontSize: 11.5,
                color: WARN,
              }}
            >
              전입 이전 학기 이수 과목이 입력되지 않았습니다.
            </div>
          ),
        )}
      {((_b = result.extraEntries) === null || _b === void 0 ? void 0 : _b.length) > 0 &&
        section(
          "공동교육과정 등 추가 이수 과목",
          table(
            ["구분", "과목명", "교과(군)", "학기", "학점"],
            result.extraEntries.map((x) => [
              x.kind,
              x.name,
              x.group,
              x.semester ? formatSemester(x.semester) : "-",
              x.credit,
            ]),
          ),
        )}
      {otherTaken.length > 0 &&
        section(
          "그 밖에 학점에 반영된 과목 (선택그룹 밖 · 편제표와 학기 다름 등)",
          table(
            ["학기", "교과(군)", "과목명", "학점"],
            otherTaken.map((t) => [t.semester ? formatSemester(t.semester) : "-", t.group, t.name, t.credit]),
          ),
        )}
      {(((_d = (_c = result.change) === null || _c === void 0 ? void 0 : _c.added) === null || _d === void 0
        ? void 0
        : _d.length) > 0 ||
        ((_f = (_e = result.change) === null || _e === void 0 ? void 0 : _e.removed) === null || _f === void 0
          ? void 0
          : _f.length) > 0) &&
        section(
          "상담 후 변경한 선택과목",
          <div
            style={{
              fontSize: 11.5,
              lineHeight: 1.7,
            }}
          >
            {result.change.added.map((e, i) => (
              <div key={`a${i}`}>
                {"+ \uCD94\uAC00 \u00B7 "}
                {e.name}
                {" ("}
                {formatSemester(e.semester)})
              </div>
            ))}
            {result.change.removed.map((e, i) => (
              <div key={`r${i}`}>
                {"\u2212 \uCDE8\uC18C \u00B7 "}
                {e.name}
                {" ("}
                {formatSemester(e.semester)})
              </div>
            ))}
            <div
              style={{
                marginTop: 4,
              }}
            >
              {"\uC218\uAC15\uC2E0\uCCAD \uC2DC\uC2A4\uD15C \uBC18\uC601: "}
              {result.change.applyState === "applied"
                ? `반영함 (${formatDateYmd(new Date(result.change.appliedAt))})`
                : "반영 확인 전"}
            </div>
          </div>,
        )}
      <div
        style={{
          marginTop: 28,
          borderTop: `2px solid ${INK}`,
          paddingTop: 14,
          breakInside: "avoid",
        }}
      >
        <div
          style={{
            fontSize: 13,
          }}
        >
          위 과목들을 선택하였음을 확인합니다.
        </div>
        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            gap: 36,
            marginTop: 20,
            fontSize: 13,
          }}
        >
          <div>
            {"\uB2F4\uB2F9\uAD50\uC0AC: "}
            {teacherName || "        "}
            {" \u00A0 (\uC11C\uBA85) ______________"}
          </div>
          <div>
            {"\uD559\uC0DD: "}
            {result.name}
            {" \u00A0 (\uC11C\uBA85) ______________"}
          </div>
        </div>
      </div>
    </div>
  );
}
// Printable summary for one or more students. Always mounted (hidden on screen), and shown via a
// plain @media print rule using normal block flow — NOT position:absolute — so multi-page /
// multi-student content paginates correctly instead of overlapping across page breaks.
function PrintRoot({
  results,
  groups,
  priorityGroups,
  reqOverride,
  curriculumFile,
  classTeachers,
  qrSheet,
  listSheet,
}) {
  if (listSheet)
    return (
      <div id="print-root">
        <ListPrintPage sheet={listSheet} />
      </div>
    );
  if (qrSheet && qrSheet.items.length > 0)
    return (
      <div id="print-root">
        <SurveyQrPrintPages sheet={qrSheet} />
      </div>
    );
  if (!results || results.length === 0) return <div id="print-root" />;
  return (
    <div id="print-root">
      {results.map((r, i) => (
        <StudentPrintCard
          key={r.key}
          result={r}
          groups={groups}
          priorityGroups={priorityGroups}
          reqOverride={reqOverride}
          curriculumFile={curriculumFile}
          classTeachers={classTeachers}
          pageBreak={i < results.length - 1}
        />
      ))}
    </div>
  );
}
// ================= PAGE 1: 편제표 =================
function CurriculumPage({
  curriculumFile,
  curriculumError,
  handleCurriculum,
  clearCurriculum,
  minCredit,
  setMinCredit,
  reqOverride,
  setReqOverride,
  showReq,
  setShowReq,
  groups,
  onNext,
  onGoBuilder,
  priorityGroups,
  setPriorityGroups,
  admissionYear,
  setAdmissionYear,
  creativeActivityCredit,
  setCreativeActivityCredit,
  curriculumIssues,
}) {
  return (
    <div>
      <div
        style={{
          marginBottom: 20,
        }}
      >
        <div
          style={{
            fontSize: 13.5,
            color: MUTED,
          }}
        >
          교육과정편제표 엑셀 파일(편제표입력 시트 포함)을 올려주세요. 교과영역별 필수이수학점을 자동으로 읽어옵니다.
          엑셀 양식을 쓰지 않고 화면에서 클릭으로 만들고 싶다면{" "}
          <button
            onClick={onGoBuilder}
            style={{
              border: "none",
              background: "none",
              color: ACCENT,
              fontWeight: 700,
              cursor: "pointer",
              padding: 0,
              fontSize: 13.5,
              textDecoration: "underline",
            }}
          >
            편제표 직접 작성 (수작업)
          </button>{" "}
          탭을 이용하세요. 둘 중 한 가지만 하면 되고, 나중에 적용한 편제표가 사용됩니다.
        </div>
      </div>
      <UploadCard
        title="편제표 업로드"
        desc="편제표입력 시트가 포함된 엑셀 파일 1개. 학교 편제에 맞게 자유롭게 수정한 파일이어도 됩니다."
        multiple={false}
        onFiles={handleCurriculum}
        files={
          curriculumFile
            ? [
                {
                  name: `${curriculumFile.name} · ${curriculumFile.courses.length}과목`,
                },
              ]
            : curriculumError
              ? [
                  {
                    name: "업로드 실패",
                    status: "error",
                    error: curriculumError,
                  },
                ]
              : []
        }
        onRemove={clearCurriculum}
      />
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          marginTop: 10,
        }}
      >
        <button
          onClick={downloadTemplate}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            fontSize: 12.5,
            fontWeight: 600,
            color: ACCENT,
            background: "none",
            border: "none",
            padding: "4px 2px",
            cursor: "pointer",
          }}
        >
          <Download size={13} />
          편제표 기본 양식 다운로드
        </button>
        <span
          style={{
            fontSize: 12,
            color: MUTED,
          }}
        >
          처음 만드는 경우, 이 양식을 내려받아 ‘편제표입력’ 시트만 채워서 올려주세요. 양식의 모든 칸은 셀 서식이
          ‘텍스트’라 1-1 같은 개설학기가 날짜로 바뀌지 않습니다.
        </span>
      </div>
      {curriculumFile && (
        <div
          style={{
            background: "#FEFBEA",
            border: "1px solid #F0E2A3",
            borderRadius: 10,
            padding: "16px 18px",
            marginTop: 18,
          }}
        >
          <div
            style={{
              fontSize: 14,
              fontWeight: 700,
              marginBottom: 6,
              display: "flex",
              alignItems: "center",
              gap: 6,
            }}
          >
            ⭐ 중점 관리 교과(군)
          </div>
          <div
            style={{
              fontSize: 12,
              color: MUTED,
              marginBottom: 10,
              lineHeight: 1.5,
            }}
          >
            학교마다 잘 안 채워지는 교과군이 정해져 있습니다. 여기서 고른 교과군은 ③ 졸업요건 점검에서 맨 앞 열로 오고,
            표에서 노란 음영으로 강조되며, 인쇄되는 확인서에서도 먼저 나옵니다.
          </div>
          <div
            style={{
              display: "flex",
              flexWrap: "wrap",
              gap: "6px 18px",
              marginBottom: 12,
            }}
          >
            {groups.map((g) => {
              var _a;
              return (
                <label
                  key={g}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 5,
                    fontSize: 13,
                    cursor: "pointer",
                  }}
                >
                  <input
                    type="checkbox"
                    checked={!!priorityGroups[g]}
                    onChange={(e) =>
                      setPriorityGroups((prev) => ({
                        ...prev,
                        [g]: e.target.checked,
                      }))
                    }
                  />
                  {g}
                  {" ("}
                  {(_a = reqOverride[g]) !== null && _a !== void 0 ? _a : 0})
                </label>
              );
            })}
          </div>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
              gap: 14,
              marginTop: 14,
            }}
          >
            <div>
              <label
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  display: "block",
                  marginBottom: 4,
                }}
              >
                입학연도
              </label>
              <input
                type="number"
                value={admissionYear}
                onChange={(e) => setAdmissionYear(Number(e.target.value) || admissionYear)}
                style={{
                  width: 100,
                  padding: "7px 9px",
                  borderRadius: 6,
                  border: `1px solid ${LINE}`,
                  fontSize: 13,
                }}
              />
            </div>
            <div>
              <label
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  display: "block",
                  marginBottom: 4,
                }}
              >
                창의적 체험활동(창체) 학점
              </label>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                <input
                  type="number"
                  value={creativeActivityCredit}
                  onChange={(e) => setCreativeActivityCredit(Number(e.target.value) || 0)}
                  style={{
                    width: 80,
                    padding: "7px 9px",
                    borderRadius: 6,
                    border: `1px solid ${LINE}`,
                    fontSize: 13,
                  }}
                />
                <span
                  style={{
                    fontSize: 12,
                    color: MUTED,
                  }}
                >
                  학점
                </span>
              </div>
            </div>
            <div>
              <label
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  display: "block",
                  marginBottom: 4,
                }}
              >
                교과 이수학점 기준(졸업)
              </label>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                <input
                  type="number"
                  value={minCredit}
                  onChange={(e) => setMinCredit(Number(e.target.value) || 0)}
                  style={{
                    width: 80,
                    padding: "7px 9px",
                    borderRadius: 6,
                    border: `1px solid ${LINE}`,
                    fontSize: 13,
                  }}
                />
                <span
                  style={{
                    fontSize: 12,
                    color: MUTED,
                  }}
                >
                  학점
                </span>
              </div>
            </div>
          </div>
          <div
            style={{
              fontSize: 12,
              color: MUTED,
              marginTop: 10,
            }}
          >
            {"\uAD50\uACFC "}
            {minCredit}
            {" + \uCC3D\uCCB4 "}
            {creativeActivityCredit}
            {" = \uCD1D "}
            <b
              style={{
                color: INK,
              }}
            >
              {minCredit + creativeActivityCredit}
            </b>
            학점 · 교과영역별 필수이수학점은 편제표에서 자동으로 읽어와 계산에 사용됩니다 ({groups.length}개 영역).
          </div>
        </div>
      )}
      {curriculumFile && <CurriculumIssues issues={curriculumIssues || []} />}
      {curriculumFile && <CourseListBySemester courses={curriculumFile.courses} />}
      {curriculumFile && (
        <div
          style={{
            display: "flex",
            justifyContent: "flex-end",
            marginTop: 20,
          }}
        >
          <button
            onClick={onNext}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              fontSize: 13.5,
              fontWeight: 700,
              color: "#fff",
              background: ACCENT,
              border: "none",
              borderRadius: 8,
              padding: "10px 18px",
              cursor: "pointer",
            }}
          >
            {"\uB2E4\uC74C: \uC218\uAC15\uC2E0\uCCAD \uD30C\uC77C \uC5C5\uB85C\uB4DC "}
            <ArrowRight size={15} />
          </button>
        </div>
      )}
    </div>
  );
}
function CourseListBySemester({ courses }) {
  const [openSem, setOpenSem] = useState(null);
  const bySem = useMemo(() => {
    const m = new Map();
    courses.forEach((c) => {
      const sem = c.semester || "학기 미상";
      if (!m.has(sem))
        m.set(sem, {
          mandatory: [],
          elective: [],
        });
      const b = m.get(sem);
      if (c.gubun === "학교지정") b.mandatory.push(c);
      else b.elective.push(c);
    });
    return Array.from(m.entries())
      .map(([sem, v]) => ({
        sem,
        label: formatSemester(sem),
        mandatory: sortCoursesForList(v.mandatory, courses),
        elective: sortCoursesForList(v.elective, courses),
      }))
      .sort((a, b) => {
        const ka = semesterSortKey(a.sem);
        const kb = semesterSortKey(b.sem);
        return ka[0] - kb[0] || ka[1] - kb[1];
      });
  }, [courses]);
  return (
    <div
      style={{
        marginTop: 16,
      }}
    >
      <div
        style={{
          fontSize: 13.5,
          fontWeight: 700,
          marginBottom: 10,
        }}
      >
        학기별 과목 목록{" "}
        <span
          style={{
            fontWeight: 400,
            fontSize: 12,
            color: MUTED,
          }}
        >
          · 학교지정·학생선택 안에서 학점이 높은 과목부터, 같은 학점은 교과(군) 순
        </span>
      </div>
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 8,
        }}
      >
        {bySem.map((s) => {
          const open = openSem === s.sem;
          return (
            <div
              key={s.sem}
              style={{
                background: "#fff",
                border: `1px solid ${LINE}`,
                borderRadius: 10,
              }}
            >
              <button
                onClick={() => setOpenSem(open ? null : s.sem)}
                style={{
                  width: "100%",
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  padding: "12px 16px",
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                  fontSize: 13.5,
                  fontWeight: 700,
                  color: INK,
                }}
              >
                <span>
                  {s.label}{" "}
                  <span
                    style={{
                      fontWeight: 400,
                      fontSize: 12,
                      color: MUTED,
                    }}
                  >
                    {"\u00B7 \uD559\uAD50\uC9C0\uC815 "}
                    {s.mandatory.length}
                    {" \u00B7 \uD559\uC0DD\uC120\uD0DD "}
                    {s.elective.length}
                  </span>
                </span>
                {open ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
              </button>
              {open && (
                <div
                  style={{
                    padding: "0 16px 16px",
                    overflowX: "auto",
                  }}
                >
                  <table
                    style={{
                      borderCollapse: "collapse",
                      width: "100%",
                      fontSize: 12.5,
                      minWidth: 480,
                    }}
                  >
                    <thead>
                      <tr
                        style={{
                          background: PAPER,
                        }}
                      >
                        <th style={thStyle}>구분</th>
                        <th style={thStyle}>교과(군)</th>
                        <th style={thStyle}>과목명</th>
                        <th
                          style={{
                            ...thStyle,
                            textAlign: "right",
                          }}
                        >
                          학점
                        </th>
                        <th style={thStyle}>선택구분</th>
                      </tr>
                    </thead>
                    <tbody>
                      {[...s.mandatory, ...s.elective].length === 0 && (
                        <tr>
                          <td
                            colSpan={5}
                            style={{
                              padding: "12px",
                              color: MUTED,
                              fontSize: 12.5,
                            }}
                          >
                            등록된 과목이 없습니다.
                          </td>
                        </tr>
                      )}
                      {s.mandatory.map((c, i) => (
                        <tr
                          key={`m-${i}`}
                          style={{
                            borderTop: `1px solid ${LINE}`,
                          }}
                        >
                          <td style={tdStyle}>
                            <span
                              style={{
                                fontSize: 11,
                                fontWeight: 700,
                                color: OK,
                                background: OK_BG,
                                borderRadius: 999,
                                padding: "2px 8px",
                              }}
                            >
                              학교지정
                            </span>
                          </td>
                          <td style={tdStyle}>{c.group}</td>
                          <td style={tdStyle}>{c.name}</td>
                          <td
                            style={{
                              ...tdStyle,
                              textAlign: "right",
                            }}
                          >
                            {c.credit}
                          </td>
                          <td
                            style={{
                              ...tdStyle,
                              color: MUTED,
                            }}
                          >
                            -
                          </td>
                        </tr>
                      ))}
                      {s.elective.map((c, i) => (
                        <tr
                          key={`e-${i}`}
                          style={{
                            borderTop: `1px solid ${LINE}`,
                          }}
                        >
                          <td style={tdStyle}>
                            <span
                              style={{
                                fontSize: 11,
                                fontWeight: 700,
                                color: ACCENT,
                                background: ACCENT_BG,
                                borderRadius: 999,
                                padding: "2px 8px",
                              }}
                            >
                              학생선택
                            </span>
                          </td>
                          <td style={tdStyle}>{c.group}</td>
                          <td style={tdStyle}>{c.name}</td>
                          <td
                            style={{
                              ...tdStyle,
                              textAlign: "right",
                            }}
                          >
                            {c.credit}
                          </td>
                          <td
                            style={{
                              ...tdStyle,
                              color: ACCENT,
                              fontWeight: 600,
                            }}
                          >
                            {c.selectGroup || "-"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}
const thStyle = {
  textAlign: "left",
  padding: "8px 10px",
  fontSize: 11.5,
  fontWeight: 700,
  color: MUTED,
  whiteSpace: "nowrap",
};
const tdStyle = {
  padding: "7px 10px",
  whiteSpace: "nowrap",
};
// ================= BUILDER PAGE: 편제표 직접 작성 =================
function BuilderPage({ builderData, setBuilderData, onApply }) {
  const [activeSem, setActiveSem] = useState("1-1");
  const [activeBucket, setActiveBucket] = useState({
    type: "mandatory",
  });
  const [search, setSearch] = useState("");
  const [category, setCategory] = useState(null);
  const [customSubjects, setCustomSubjects] = useState([]);
  const [showCustomForm, setShowCustomForm] = useState(false);
  const [customForm, setCustomForm] = useState({
    name: "",
    group: BUILDER_CATEGORIES[0],
    type: "일반",
    credit: 4,
  });
  const [poolForm, setPoolForm] = useState({
    n: 4,
    credit: 4,
  });
  const [showPoolForm, setShowPoolForm] = useState(false);
  const palette = useMemo(() => [...NATIONAL_SUBJECT_MASTER, ...customSubjects], [customSubjects]);
  const filteredPalette = useMemo(() => {
    return palette.filter((s) => {
      if (category && s.group !== category) return false;
      if (search.trim() && !s.name.includes(search.trim())) return false;
      return true;
    });
  }, [palette, category, search]);
  const sem = builderData[activeSem];
  const isInBucket = (name) => {
    if (activeBucket.type === "mandatory") return sem.mandatory.some((c) => c.name === name);
    const pool = sem.pools.find((p) => p.id === activeBucket.poolId);
    return pool ? pool.courses.some((c) => c.name === name) : false;
  };
  const updateSem = (updater) => {
    setBuilderData((prev) => ({
      ...prev,
      [activeSem]: updater(prev[activeSem]),
    }));
  };
  const toggleSubject = (subj) => {
    if (activeBucket.type === "mandatory") {
      updateSem((s0) => {
        const exists = s0.mandatory.some((c) => c.name === subj.name);
        return {
          ...s0,
          mandatory: exists
            ? s0.mandatory.filter((c) => c.name !== subj.name)
            : [
                ...s0.mandatory,
                {
                  name: subj.name,
                  group: subj.group,
                  type: subj.type,
                  credit: subj.credit,
                },
              ],
        };
      });
    } else {
      updateSem((s0) => ({
        ...s0,
        pools: s0.pools.map((p) => {
          if (p.id !== activeBucket.poolId) return p;
          const exists = p.courses.some((c) => c.name === subj.name);
          return {
            ...p,
            courses: exists
              ? p.courses.filter((c) => c.name !== subj.name)
              : [
                  ...p.courses,
                  {
                    name: subj.name,
                    group: subj.group,
                    type: subj.type,
                  },
                ],
          };
        }),
      }));
    }
  };
  const removeFromMandatory = (name) =>
    updateSem((s0) => ({
      ...s0,
      mandatory: s0.mandatory.filter((c) => c.name !== name),
    }));
  const setMandatoryCredit = (name, credit) =>
    updateSem((s0) => ({
      ...s0,
      mandatory: s0.mandatory.map((c) =>
        c.name === name
          ? {
              ...c,
              credit,
            }
          : c,
      ),
    }));
  const removeFromPool = (poolId, name) =>
    updateSem((s0) => ({
      ...s0,
      pools: s0.pools.map((p) =>
        p.id === poolId
          ? {
              ...p,
              courses: p.courses.filter((c) => c.name !== name),
            }
          : p,
      ),
    }));
  const removePool = (poolId) => {
    updateSem((s0) => ({
      ...s0,
      pools: s0.pools.filter((p) => p.id !== poolId),
    }));
    if (activeBucket.type === "pool" && activeBucket.poolId === poolId)
      setActiveBucket({
        type: "mandatory",
      });
  };
  const addPool = () => {
    const id = `pool-${Date.now()}`;
    updateSem((s0) => ({
      ...s0,
      pools: [
        ...s0.pools,
        {
          id,
          requiredN: poolForm.n,
          creditPerCourse: poolForm.credit,
          courses: [],
        },
      ],
    }));
    setActiveBucket({
      type: "pool",
      poolId: id,
    });
    setShowPoolForm(false);
  };
  const setPoolField = (poolId, field, value) =>
    updateSem((s0) => ({
      ...s0,
      pools: s0.pools.map((p) =>
        p.id === poolId
          ? {
              ...p,
              [field]: value,
            }
          : p,
      ),
    }));
  const addCustomSubject = () => {
    if (!customForm.name.trim()) return;
    setCustomSubjects((prev) => [
      ...prev,
      {
        ...customForm,
        name: customForm.name.trim(),
        credit: Number(customForm.credit) || 4,
      },
    ]);
    setShowCustomForm(false);
    setCustomForm({
      name: "",
      group: BUILDER_CATEGORIES[0],
      type: "일반",
      credit: 4,
    });
  };
  const semesterTotal = (semCode) => {
    const s0 = builderData[semCode];
    const mandatoryTotal = s0.mandatory.reduce((a, c) => a + (Number(c.credit) || 0), 0);
    const poolTotal = s0.pools.reduce((a, p) => a + p.requiredN * p.creditPerCourse, 0);
    return mandatoryTotal + poolTotal;
  };
  const overallSummary = useMemo(() => {
    const groupTotals = {};
    let grandTotal = 0;
    SEMESTER_CODES.forEach((semCode) => {
      const s0 = builderData[semCode];
      s0.mandatory.forEach((c) => {
        groupTotals[c.group] = (groupTotals[c.group] || 0) + (Number(c.credit) || 0);
        grandTotal += Number(c.credit) || 0;
      });
      s0.pools.forEach((p) => {
        const poolTotal = p.requiredN * p.creditPerCourse;
        grandTotal += poolTotal;
        if (p.courses.length === 0) return;
        const counts = {};
        p.courses.forEach((c) => (counts[c.group] = (counts[c.group] || 0) + 1));
        Object.entries(counts).forEach(([g, cnt]) => {
          groupTotals[g] = (groupTotals[g] || 0) + (poolTotal * cnt) / p.courses.length;
        });
      });
    });
    return {
      groupTotals,
      grandTotal,
    };
  }, [builderData]);
  const displayGroups = useMemo(() => {
    const merged = mergeLifestyleGroups(
      Object.keys(overallSummary.groupTotals).map((g) => ({
        group: g,
      })),
      {},
      Object.keys(overallSummary.groupTotals),
    );
    // fold groupTotals through the same merge mapping for display
    const totals = {};
    Object.entries(overallSummary.groupTotals).forEach(([g, v]) => {
      const mapped = toCurriculumGroup(g);
      totals[mapped] = (totals[mapped] || 0) + v;
    });
    return {
      groups: merged.groups,
      totals,
    };
  }, [overallSummary]);
  const activePoolObj = activeBucket.type === "pool" ? sem.pools.find((p) => p.id === activeBucket.poolId) : null;
  return (
    <div>
      <div
        style={{
          background: "#fff",
          border: `1px solid ${LINE}`,
          borderRadius: 10,
          padding: "12px 16px",
          fontSize: 12.5,
          color: MUTED,
          marginBottom: 16,
          lineHeight: 1.6,
        }}
      >
        <b
          style={{
            color: INK,
          }}
        >
          엑셀 양식 없이 화면에서 편제표를 만드는 곳입니다.
        </b>
        {
          " \u2018\uD3B8\uC81C\uD45C \uC5C5\uB85C\uB4DC\u2019\uB85C \uC62C\uB9B0 \uD30C\uC77C\uC774 \uC788\uC73C\uBA74 \uADF8 \uB0B4\uC6A9\uC774 \uBBF8\uB9AC \uCC44\uC6CC\uC838 \uC788\uC5B4 \uACE0\uCCD0 \uC4F8 \uC218 \uC788\uC2B5\uB2C8\uB2E4."
        }{" "}
        {
          '\uC67C\uCABD\uC5D0\uC11C \uACFC\uBAA9\uC744 \uB204\uB974\uBA74 \uC624\uB978\uCABD \uC120\uD0DD\uB41C \uCE78\uC5D0 \uB2F4\uAE41\uB2C8\uB2E4. \uACFC\uBAA9\uBA85\uC740 \uC9C1\uC811 \uCE58\uC9C0 \uC54A\uC73C\uB2C8 \uC624\uD0C0\uAC00 \uB098\uC9C0 \uC54A\uACE0, \uD559\uC810\uC740 \uADF8\uB8F9\uC774 \uC815\uD569\uB2C8\uB2E4 \u2014 \uD0DDN \uADF8\uB8F9\uC758 "\uACFC\uBAA9\uB2F9 \uD559\uC810"\uC744 \uD55C \uBC88\uB9CC \uB123\uC73C\uBA74 \uADF8 \uADF8\uB8F9 \uC804\uCCB4\uC5D0 \uC801\uC6A9\uB429\uB2C8\uB2E4. \uBAA9\uB85D\uC5D0 \uC5C6\uB294 \uACFC\uBAA9\uC740'
        }{" "}
        <b
          style={{
            color: INK,
          }}
        >
          + 고시 외 과목
        </b>
        으로 직접 추가할 수 있어요.
      </div>
      <div
        style={{
          display: "flex",
          gap: 6,
          marginBottom: 16,
          flexWrap: "wrap",
        }}
      >
        {SEMESTER_CODES.map((s0) => (
          <button
            key={s0}
            onClick={() => {
              setActiveSem(s0);
              setActiveBucket({
                type: "mandatory",
              });
            }}
            style={{
              border: `1px solid ${activeSem === s0 ? ACCENT : LINE}`,
              background: activeSem === s0 ? ACCENT : "#fff",
              color: activeSem === s0 ? "#fff" : INK,
              borderRadius: 999,
              padding: "7px 14px",
              fontSize: 12.5,
              fontWeight: 700,
              cursor: "pointer",
            }}
          >
            {formatSemester(s0)} {semesterTotal(s0)}
          </button>
        ))}
      </div>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "1.1fr 1fr",
          gap: 16,
        }}
      >
        <div
          style={{
            background: "#fff",
            border: `1px solid ${LINE}`,
            borderRadius: 10,
            padding: 14,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 8,
            }}
          >
            과목 팔레트 (누르면 담김)
          </div>
          <div
            style={{
              display: "flex",
              gap: 8,
              marginBottom: 10,
            }}
          >
            <div
              style={{
                position: "relative",
                flex: 1,
              }}
            >
              <Search
                size={13}
                style={{
                  position: "absolute",
                  left: 8,
                  top: 9,
                  color: MUTED,
                }}
              />
              <input
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                placeholder="과목명 검색 (예: 미적분)"
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  padding: "7px 9px 7px 26px",
                  borderRadius: 7,
                  border: `1px solid ${LINE}`,
                  fontSize: 12.5,
                }}
              />
            </div>
            <button
              onClick={() => setShowCustomForm((v) => !v)}
              style={{
                fontSize: 12,
                fontWeight: 700,
                color: ACCENT,
                background: ACCENT_BG,
                border: "none",
                borderRadius: 7,
                padding: "0 10px",
                cursor: "pointer",
                whiteSpace: "nowrap",
              }}
            >
              + 고시 외 과목
            </button>
          </div>
          {showCustomForm && (
            <div
              style={{
                background: PAPER,
                borderRadius: 8,
                padding: 10,
                marginBottom: 10,
                display: "flex",
                flexDirection: "column",
                gap: 6,
              }}
            >
              <input
                value={customForm.name}
                onChange={(e) =>
                  setCustomForm((f) => ({
                    ...f,
                    name: e.target.value,
                  }))
                }
                placeholder="과목명"
                style={{
                  padding: "6px 8px",
                  borderRadius: 6,
                  border: `1px solid ${LINE}`,
                  fontSize: 12.5,
                }}
              />
              <div
                style={{
                  display: "flex",
                  gap: 6,
                }}
              >
                <select
                  value={customForm.group}
                  onChange={(e) =>
                    setCustomForm((f) => ({
                      ...f,
                      group: e.target.value,
                    }))
                  }
                  style={{
                    flex: 1,
                    padding: "6px 8px",
                    borderRadius: 6,
                    border: `1px solid ${LINE}`,
                    fontSize: 12.5,
                  }}
                >
                  {BUILDER_CATEGORIES.map((c) => (
                    <option key={c} value={c}>
                      {c}
                    </option>
                  ))}
                </select>
                <select
                  value={customForm.type}
                  onChange={(e) =>
                    setCustomForm((f) => ({
                      ...f,
                      type: e.target.value,
                    }))
                  }
                  style={{
                    width: 90,
                    padding: "6px 8px",
                    borderRadius: 6,
                    border: `1px solid ${LINE}`,
                    fontSize: 12.5,
                  }}
                >
                  {["공통", "일반", "진로", "융합"].map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
                <input
                  type="number"
                  value={customForm.credit}
                  onChange={(e) =>
                    setCustomForm((f) => ({
                      ...f,
                      credit: e.target.value,
                    }))
                  }
                  style={{
                    width: 60,
                    padding: "6px 8px",
                    borderRadius: 6,
                    border: `1px solid ${LINE}`,
                    fontSize: 12.5,
                  }}
                />
              </div>
              <button
                onClick={addCustomSubject}
                style={{
                  fontSize: 12,
                  fontWeight: 700,
                  color: "#fff",
                  background: ACCENT,
                  border: "none",
                  borderRadius: 6,
                  padding: "6px 0",
                  cursor: "pointer",
                }}
              >
                팔레트에 추가
              </button>
            </div>
          )}
          <div
            style={{
              display: "flex",
              gap: 5,
              flexWrap: "wrap",
              marginBottom: 10,
            }}
          >
            <button
              onClick={() => setCategory(null)}
              style={{
                fontSize: 11.5,
                fontWeight: 600,
                padding: "3px 9px",
                borderRadius: 999,
                border: `1px solid ${!category ? ACCENT : LINE}`,
                background: !category ? ACCENT_BG : "#fff",
                color: !category ? ACCENT : MUTED,
                cursor: "pointer",
              }}
            >
              전체
            </button>
            {BUILDER_CATEGORIES.map((c) => (
              <button
                key={c}
                onClick={() => setCategory(c)}
                style={{
                  fontSize: 11.5,
                  fontWeight: 600,
                  padding: "3px 9px",
                  borderRadius: 999,
                  border: `1px solid ${category === c ? ACCENT : LINE}`,
                  background: category === c ? ACCENT_BG : "#fff",
                  color: category === c ? ACCENT : MUTED,
                  cursor: "pointer",
                }}
              >
                {c}
              </button>
            ))}
          </div>
          <div
            style={{
              maxHeight: 420,
              overflowY: "auto",
              display: "flex",
              flexDirection: "column",
              gap: 2,
            }}
          >
            {filteredPalette.map((s0, i) => {
              const on = isInBucket(s0.name);
              return (
                <div
                  key={i}
                  onClick={() => toggleSubject(s0)}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "space-between",
                    padding: "7px 8px",
                    borderRadius: 6,
                    cursor: "pointer",
                    background: on ? ACCENT_BG : "transparent",
                    fontSize: 12.5,
                  }}
                >
                  <span
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 6,
                    }}
                  >
                    <span
                      style={{
                        color: on ? ACCENT : "transparent",
                        fontWeight: 700,
                        width: 12,
                      }}
                    >
                      {on ? "✓" : "+"}
                    </span>
                    <span
                      style={{
                        color: on ? ACCENT : INK,
                        fontWeight: on ? 700 : 400,
                      }}
                    >
                      {s0.name}
                    </span>
                  </span>
                  <span
                    style={{
                      color: MUTED,
                      fontSize: 11,
                    }}
                  >
                    {s0.group} {s0.credit}
                  </span>
                </div>
              );
            })}
            {filteredPalette.length === 0 && (
              <div
                style={{
                  fontSize: 12,
                  color: MUTED,
                  padding: 8,
                }}
              >
                검색 결과가 없습니다.
              </div>
            )}
          </div>
        </div>
        <div
          style={{
            background: "#fff",
            border: `1px solid ${LINE}`,
            borderRadius: 10,
            padding: 14,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 2,
            }}
          >
            {formatSemester(activeSem)}
            {" \uD3B8\uC131"}
          </div>
          <div
            style={{
              fontSize: 11.5,
              color: MUTED,
              marginBottom: 10,
            }}
          >
            담을 곳을 먼저 고르고, 왼쪽에서 과목을 누르세요.
          </div>
          <div
            onClick={() =>
              setActiveBucket({
                type: "mandatory",
              })
            }
            style={{
              border: `1.5px solid ${activeBucket.type === "mandatory" ? ACCENT : LINE}`,
              borderRadius: 9,
              padding: 10,
              marginBottom: 10,
              cursor: "pointer",
              background: activeBucket.type === "mandatory" ? ACCENT_BG : PAPER,
            }}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                marginBottom: 6,
              }}
            >
              <span
                style={{
                  fontSize: 12.5,
                  fontWeight: 700,
                  color: OK,
                }}
              >
                ● 학교지정
              </span>
              <span
                style={{
                  fontSize: 11,
                  color: MUTED,
                }}
              >
                {sem.mandatory.length}
                {"\uACFC\uBAA9 \uAC1C\uC124 \u00B7 "}
                {sem.mandatory.reduce((a, c) => a + (Number(c.credit) || 0), 0)}학점
              </span>
            </div>
            <div
              style={{
                display: "flex",
                flexWrap: "wrap",
                gap: 6,
              }}
            >
              {sem.mandatory.map((c) => (
                <div
                  key={c.name}
                  onClick={(e) => e.stopPropagation()}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 4,
                    background: "#fff",
                    border: `1px solid ${LINE}`,
                    borderRadius: 7,
                    padding: "4px 6px",
                    fontSize: 12,
                  }}
                >
                  <span>{c.name}</span>
                  <input
                    type="number"
                    value={c.credit}
                    onChange={(e) => setMandatoryCredit(c.name, Number(e.target.value) || 0)}
                    style={{
                      width: 34,
                      border: `1px solid ${LINE}`,
                      borderRadius: 4,
                      fontSize: 11.5,
                      textAlign: "center",
                    }}
                  />
                  <button
                    onClick={() => removeFromMandatory(c.name)}
                    style={{
                      border: "none",
                      background: "none",
                      cursor: "pointer",
                      color: MUTED,
                      padding: 0,
                    }}
                  >
                    <X size={11} />
                  </button>
                </div>
              ))}
              {sem.mandatory.length === 0 && (
                <span
                  style={{
                    fontSize: 11.5,
                    color: MUTED,
                  }}
                >
                  없음
                </span>
              )}
            </div>
          </div>
          {sem.pools.map((p) => (
            <div
              key={p.id}
              onClick={() =>
                setActiveBucket({
                  type: "pool",
                  poolId: p.id,
                })
              }
              style={{
                border: `1.5px solid ${activeBucket.type === "pool" && activeBucket.poolId === p.id ? ACCENT : LINE}`,
                borderRadius: 9,
                padding: 10,
                marginBottom: 10,
                cursor: "pointer",
                background: activeBucket.type === "pool" && activeBucket.poolId === p.id ? ACCENT_BG : PAPER,
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "space-between",
                  marginBottom: 6,
                  flexWrap: "wrap",
                  gap: 6,
                }}
              >
                <span
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 6,
                    fontSize: 12.5,
                    fontWeight: 700,
                    color: ACCENT,
                  }}
                  onClick={(e) => e.stopPropagation()}
                >
                  ○ 택
                  <input
                    type="number"
                    value={p.requiredN}
                    onChange={(e) => setPoolField(p.id, "requiredN", Number(e.target.value) || 0)}
                    style={{
                      width: 34,
                      border: `1px solid ${LINE}`,
                      borderRadius: 4,
                      fontSize: 12,
                      textAlign: "center",
                    }}
                  />
                  과목당
                  <input
                    type="number"
                    value={p.creditPerCourse}
                    onChange={(e) => setPoolField(p.id, "creditPerCourse", Number(e.target.value) || 0)}
                    style={{
                      width: 34,
                      border: `1px solid ${LINE}`,
                      borderRadius: 4,
                      fontSize: 12,
                      textAlign: "center",
                    }}
                  />
                  {"\uD559\uC810 (\uCD1D "}
                  {p.requiredN * p.creditPerCourse}학점)
                </span>
                <button
                  onClick={(e) => {
                    e.stopPropagation();
                    removePool(p.id);
                  }}
                  style={{
                    border: "none",
                    background: "none",
                    cursor: "pointer",
                    color: WARN,
                    fontSize: 11,
                  }}
                >
                  그룹 삭제
                </button>
              </div>
              <div
                style={{
                  display: "flex",
                  flexWrap: "wrap",
                  gap: 6,
                }}
              >
                {p.courses.map((c) => (
                  <div
                    key={c.name}
                    onClick={(e) => e.stopPropagation()}
                    style={{
                      display: "flex",
                      alignItems: "center",
                      gap: 4,
                      background: "#fff",
                      border: `1px solid ${LINE}`,
                      borderRadius: 7,
                      padding: "4px 6px",
                      fontSize: 12,
                    }}
                  >
                    <span
                      style={{
                        color: MUTED,
                      }}
                    >
                      [{c.group}]
                    </span>
                    <span>{c.name}</span>
                    <button
                      onClick={() => removeFromPool(p.id, c.name)}
                      style={{
                        border: "none",
                        background: "none",
                        cursor: "pointer",
                        color: MUTED,
                        padding: 0,
                      }}
                    >
                      <X size={11} />
                    </button>
                  </div>
                ))}
                {p.courses.length === 0 && (
                  <span
                    style={{
                      fontSize: 11.5,
                      color: MUTED,
                    }}
                  >
                    없음 — 왼쪽에서 과목을 눌러 담아주세요
                  </span>
                )}
              </div>
            </div>
          ))}
          {showPoolForm ? (
            <div
              style={{
                display: "flex",
                alignItems: "center",
                gap: 8,
                background: PAPER,
                borderRadius: 8,
                padding: 10,
                marginBottom: 10,
              }}
            >
              <span
                style={{
                  fontSize: 12,
                }}
              >
                택
              </span>
              <input
                type="number"
                value={poolForm.n}
                onChange={(e) =>
                  setPoolForm((f) => ({
                    ...f,
                    n: Number(e.target.value) || 0,
                  }))
                }
                style={{
                  width: 44,
                  padding: "5px 6px",
                  borderRadius: 5,
                  border: `1px solid ${LINE}`,
                }}
              />
              <span
                style={{
                  fontSize: 12,
                }}
              >
                과목, 과목당
              </span>
              <input
                type="number"
                value={poolForm.credit}
                onChange={(e) =>
                  setPoolForm((f) => ({
                    ...f,
                    credit: Number(e.target.value) || 0,
                  }))
                }
                style={{
                  width: 44,
                  padding: "5px 6px",
                  borderRadius: 5,
                  border: `1px solid ${LINE}`,
                }}
              />
              <span
                style={{
                  fontSize: 12,
                }}
              >
                학점
              </span>
              <button
                onClick={addPool}
                style={{
                  marginLeft: "auto",
                  fontSize: 12,
                  fontWeight: 700,
                  color: "#fff",
                  background: ACCENT,
                  border: "none",
                  borderRadius: 6,
                  padding: "6px 12px",
                  cursor: "pointer",
                }}
              >
                추가
              </button>
              <button
                onClick={() => setShowPoolForm(false)}
                style={{
                  fontSize: 12,
                  color: MUTED,
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                }}
              >
                취소
              </button>
            </div>
          ) : (
            <button
              onClick={() => setShowPoolForm(true)}
              style={{
                width: "100%",
                border: `1.5px dashed ${LINE}`,
                borderRadius: 9,
                padding: "10px 0",
                background: "none",
                color: ACCENT,
                fontSize: 12.5,
                fontWeight: 700,
                cursor: "pointer",
                marginBottom: 10,
              }}
            >
              + 선택그룹(택N) 추가
            </button>
          )}
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              borderTop: `1px solid ${LINE}`,
              paddingTop: 10,
            }}
          >
            {"\uC774 \uD559\uAE30 "}
            <b
              style={{
                color: INK,
              }}
            >
              {semesterTotal(activeSem)}학점
            </b>
            {" (\uC9C0\uC815 "}
            {sem.mandatory.reduce((a, c) => a + (Number(c.credit) || 0), 0)}
            {" + \uC120\uD0DD"} {sem.pools.reduce((a, p) => a + p.requiredN * p.creditPerCourse, 0)})
          </div>
        </div>
      </div>
      <div
        style={{
          background: "#fff",
          border: `1px solid ${LINE}`,
          borderRadius: 10,
          padding: 16,
          marginTop: 16,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            marginBottom: 10,
          }}
        >
          <div
            style={{
              fontSize: 13.5,
              fontWeight: 700,
            }}
          >
            전체 진행 (편제표대로 이수했을 때)
          </div>
          <div
            style={{
              fontSize: 13,
            }}
          >
            {"\uCD1D "}
            <b
              style={{
                fontSize: 16,
              }}
            >
              {Math.round(overallSummary.grandTotal * 10) / 10}
            </b>
            {" / 174\uD559\uC810"}
          </div>
        </div>
        <div
          style={{
            height: 8,
            borderRadius: 999,
            background: PAPER,
            overflow: "hidden",
            marginBottom: 12,
          }}
        >
          <div
            style={{
              height: "100%",
              width: `${Math.min(100, (overallSummary.grandTotal / 174) * 100)}%`,
              background: overallSummary.grandTotal >= 174 ? OK : ACCENT,
            }}
          />
        </div>
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(140px, 1fr))",
            gap: 8,
            marginBottom: 10,
          }}
        >
          {displayGroups.groups.map((g) => {
            var _a;
            const val = Math.round((displayGroups.totals[g] || 0) * 10) / 10;
            const req =
              (_a = NATIONAL_REQUIRED_CREDITS[g]) !== null && _a !== void 0
                ? _a
                : g === LIFESTYLE_MERGE_NAME
                  ? LIFESTYLE_MERGE_CREDIT
                  : 0;
            const ok = val >= req;
            return (
              <div
                key={g}
                style={{
                  background: ok ? OK_BG : WARN_BG,
                  borderRadius: 8,
                  padding: "8px 10px",
                }}
              >
                <div
                  style={{
                    fontSize: 11.5,
                    fontWeight: 700,
                    color: ok ? OK : WARN,
                  }}
                >
                  {g}
                </div>
                <div
                  style={{
                    fontSize: 13,
                    fontWeight: 700,
                    color: ok ? OK : WARN,
                  }}
                >
                  {val}
                  {" / "}
                  {req}
                </div>
              </div>
            );
          })}
        </div>
        <div
          style={{
            fontSize: 11.5,
            color: MUTED,
            marginBottom: 14,
          }}
        >
          선택그룹은 학생마다 고르는 과목이 달라, 그룹 안 교과군 분포에 비례해 배분한 참고값입니다. 실제 판정은 ③
          졸업요건 점검에서 학생별로 합니다.
        </div>
        <div
          style={{
            display: "flex",
            gap: 8,
            flexWrap: "wrap",
          }}
        >
          <button
            onClick={() => exportBuilderToExcel(builderData)}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              fontSize: 12.5,
              fontWeight: 700,
              color: ACCENT,
              background: "#fff",
              border: `1px solid ${ACCENT}`,
              borderRadius: 7,
              padding: "8px 14px",
              cursor: "pointer",
            }}
          >
            <Download size={13} />
            {" \uC785\uB825\uC6A9 \uC5D1\uC140\uB85C \uB0B4\uBCF4\uB0B4\uAE30"}
          </button>
          <button
            onClick={onApply}
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              fontSize: 12.5,
              fontWeight: 700,
              color: "#fff",
              background: ACCENT,
              border: "none",
              borderRadius: 7,
              padding: "8px 14px",
              cursor: "pointer",
            }}
          >
            {"\uC774 \uD3B8\uC81C\uD45C\uB85C \uC9C4\uD589 "}
            <ArrowRight size={13} />
          </button>
        </div>
      </div>
    </div>
  );
}
// ================= PAGE 2: 수강신청 업로드 + 동명이인 확인 =================
function RegFileDropzone({ onFiles, label = "학기별 수강신청 파일을 끌어다 놓거나" }) {
  const inputRef = useRef(null);
  const [dragOver, setDragOver] = useState(false);
  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        const fs = Array.from(e.dataTransfer.files || []);
        if (fs.length) onFiles(fs);
      }}
      style={{
        border: `1.5px dashed ${dragOver ? ACCENT : LINE}`,
        borderRadius: 9,
        padding: "22px 16px",
        background: dragOver ? ACCENT_BG : PAPER,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: 10,
        flexWrap: "wrap",
        textAlign: "center",
      }}
    >
      <span
        style={{
          fontSize: 13,
          color: MUTED,
        }}
      >
        {label}
      </span>
      <button
        onClick={() => {
          var _a;
          return (_a = inputRef.current) === null || _a === void 0 ? void 0 : _a.click();
        }}
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          fontSize: 13,
          fontWeight: 700,
          color: "#fff",
          background: ACCENT,
          border: "none",
          borderRadius: 7,
          padding: "8px 14px",
          cursor: "pointer",
        }}
      >
        <Upload size={14} />
        {" \uD30C\uC77C \uC120\uD0DD"}
      </button>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        multiple={true}
        style={{
          display: "none",
        }}
        onChange={(e) => {
          const fs = Array.from(e.target.files || []);
          if (fs.length) onFiles(fs);
          e.target.value = "";
        }}
      />
    </div>
  );
}
function SourceTag({ source }) {
  const meta = source === SURVEY_SOURCE_ID ? SURVEY_SOURCE_META : REG_SOURCES[source];
  if (!meta) return null;
  return (
    <span
      style={{
        fontSize: 11,
        fontWeight: 700,
        color: "#fff",
        background: meta.color,
        borderRadius: 4,
        padding: "1px 6px",
        flexShrink: 0,
        whiteSpace: "nowrap",
      }}
    >
      {meta.badge} {meta.short}
    </span>
  );
}
function recordSemesterLabel(r) {
  return r.semesters && r.semesters.length > 0 ? r.semesters.map(formatSemester).join(" · ") : "학기 미상";
}
function RegFileRow({ f, info, onRemove, onRelabel, onSemester, onMove, hasCurriculum, onToggleDisabled }) {
  var _a;
  const other = f.detected ? REG_SOURCES[f.detected] : null;
  const a = info === null || info === void 0 ? void 0 : info.analysis;
  const healthy = f.status === "ok" && info && !info.needsSemester && !f.disabled;
  const originText =
    (info === null || info === void 0 ? void 0 : info.semesterOrigin) === "manual"
      ? "직접 선택"
      : (info === null || info === void 0 ? void 0 : info.semesterOrigin) === "filename"
        ? "파일 이름에서 인식"
        : (info === null || info === void 0 ? void 0 : info.semesterOrigin) === "content"
          ? `과목 구성으로 추정 (${info.contentGuess.hits}/${info.contentGuess.total}개 일치)`
          : "";
  const autoText = (info === null || info === void 0 ? void 0 : info.filenameSemester)
    ? ` → ${formatSemester(info.filenameSemester)}`
    : (info === null || info === void 0 ? void 0 : info.contentGuess)
      ? ` → ${formatSemester(info.contentGuess.sem)}`
      : " → 인식 못함";
  return (
    <div
      style={{
        background: PAPER,
        borderRadius: 8,
        padding: "8px 10px",
        fontSize: 12.5,
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
        }}
      >
        {healthy ? (
          <CheckCircle2
            size={14}
            color={OK}
            style={{
              flexShrink: 0,
            }}
          />
        ) : (
          <AlertCircle
            size={14}
            color={WARN}
            style={{
              flexShrink: 0,
            }}
          />
        )}
        <input
          value={(_a = f.label) !== null && _a !== void 0 ? _a : f.name}
          title={f.name}
          onChange={(e) => onRelabel(f.id, e.target.value)}
          style={{
            border: "none",
            background: "transparent",
            fontSize: 12.5,
            color: INK,
            flex: 1,
            minWidth: 0,
          }}
        />
        {f.count != null && (
          <span
            style={{
              color: MUTED,
              flexShrink: 0,
            }}
          >
            {"\u00B7 "}
            {f.count}명
          </span>
        )}
        {f.status === "ok" && onToggleDisabled && (
          <label
            title="끄면 이 파일은 점검에 쓰지 않습니다"
            style={{
              display: "flex",
              alignItems: "center",
              gap: 3,
              fontSize: 11.5,
              color: f.disabled ? WARN : MUTED,
              flexShrink: 0,
              cursor: "pointer",
            }}
          >
            <input type="checkbox" checked={!f.disabled} onChange={() => onToggleDisabled(f.id)} />
            계산에 사용
          </label>
        )}
        <button
          onClick={() => onRemove(f.id)}
          style={{
            border: "none",
            background: "none",
            cursor: "pointer",
            color: MUTED,
            flexShrink: 0,
          }}
        >
          <X size={14} />
        </button>
      </div>
      {f.status === "error" && (
        <div
          style={{
            color: WARN,
            marginTop: 4,
            lineHeight: 1.5,
          }}
        >
          {f.error}
        </div>
      )}
      {f.status === "ok" && f.disabled && (
        <div
          style={{
            color: WARN,
            marginTop: 4,
            lineHeight: 1.5,
          }}
        >
          ‘계산에 사용’이 꺼져 있어 이 파일은 점검에 쓰지 않습니다. 쓰려면 위의 체크를 켜세요.
        </div>
      )}
      {f.status === "wrongZone" && other && (
        <div
          style={{
            marginTop: 6,
            background: WARN_BG,
            color: WARN,
            borderRadius: 6,
            padding: "6px 8px",
            display: "flex",
            alignItems: "center",
            gap: 8,
            flexWrap: "wrap",
            lineHeight: 1.5,
          }}
        >
          <span>
            {"\uC774 \uD30C\uC77C\uC740 "}
            <b>{other.title}</b>
            {
              " \uC591\uC2DD\uC774\uB77C \uC774 \uCE78\uC5D0\uC11C\uB294 \uACC4\uC0B0\uC5D0 \uB123\uC9C0 \uC54A\uC558\uC2B5\uB2C8\uB2E4."
            }
          </span>
          <button
            onClick={() => onMove(f.id)}
            style={{
              fontSize: 12,
              fontWeight: 700,
              color: "#fff",
              background: other.color,
              border: "none",
              borderRadius: 6,
              padding: "4px 10px",
              cursor: "pointer",
            }}
          >
            {other.badge}
            {" \uCE78\uC73C\uB85C \uC62E\uAE30\uAE30"}
          </button>
        </div>
      )}
      {f.status === "ok" && info && (
        <div
          style={{
            marginTop: 6,
            display: "flex",
            flexDirection: "column",
            gap: 4,
          }}
        >
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: 6,
              flexWrap: "wrap",
            }}
          >
            <span
              style={{
                color: MUTED,
              }}
            >
              학기
            </span>
            {info.needsFileSemester ? (
              <>
                <select
                  value={f.semesterMode || "auto"}
                  onChange={(e) => onSemester(f.id, e.target.value)}
                  style={{
                    fontSize: 12,
                    border: `1px solid ${info.needsSemester ? WARN : LINE}`,
                    borderRadius: 6,
                    padding: "3px 6px",
                    color: INK,
                    background: "#fff",
                  }}
                >
                  <option value="auto">자동 인식{autoText}</option>
                  {SEMESTER_CODES.map((sc) => (
                    <option key={sc} value={sc}>
                      {formatSemester(sc)}
                    </option>
                  ))}
                </select>
                {info.effectiveSemester && (
                  <span
                    style={{
                      color: MUTED,
                    }}
                  >
                    {originText}
                  </span>
                )}
              </>
            ) : (
              info.semesterCounts.map((s) => (
                <span
                  key={s.sem}
                  style={{
                    background: "#fff",
                    border: `1px solid ${LINE}`,
                    borderRadius: 999,
                    padding: "1px 8px",
                    color: INK,
                  }}
                >
                  {formatSemester(s.sem)}
                  {" \u00B7 "}
                  {s.students}명
                </span>
              ))
            )}
          </div>
          {info.needsSemester && (
            <div
              style={{
                color: WARN,
                lineHeight: 1.5,
              }}
            >
              ⚠ 학기를 알 수 없어 계산에서 빠져 있습니다. 위에서 학기를 골라주세요.
            </div>
          )}
          {info.semesterConflict && (
            <div
              style={{
                color: WARN,
                lineHeight: 1.5,
              }}
            >
              {"\u26A0 "}
              {info.semesterOrigin === "manual" ? "직접 고른" : "파일 이름의"}
              {" \uD559\uAE30\uB294 "}
              {formatSemester(info.effectiveSemester)}인데, 과목 구성은 편제표의 {formatSemester(info.contentGuess.sem)}
              {"\uC640 "}
              {info.contentGuess.hits}/{info.contentGuess.total}개 일치합니다. 학기를 확인해주세요.
            </div>
          )}
          {a &&
            (a.semMismatch.length === 0 && a.notInCurriculum.length === 0 ? (
              <div
                style={{
                  color: OK,
                }}
              >
                {"\u2713 \uACFC\uBAA9 "}
                {a.total}개 모두 편제표와 일치 (학기까지)
              </div>
            ) : (
              <details
                style={{
                  color: WARN,
                }}
              >
                <summary
                  style={{
                    cursor: "pointer",
                    lineHeight: 1.5,
                  }}
                >
                  {"\u26A0 \uACFC\uBAA9 "}
                  {a.total}
                  {"\uAC1C \uC911 "}
                  {a.matched.length}개 일치{a.semMismatch.length > 0 && ` · 학기 다름 ${a.semMismatch.length}`}
                  {a.notInCurriculum.length > 0 && ` · 편제표에 없음 ${a.notInCurriculum.length}`}
                  {" (\uB20C\uB7EC\uC11C \uBCF4\uAE30)"}
                </summary>
                <div
                  style={{
                    marginTop: 4,
                    lineHeight: 1.6,
                    paddingLeft: 4,
                  }}
                >
                  {a.semMismatch.map((e, i) => (
                    <div key={"m" + i}>
                      {"\u00B7 "}
                      {e.name}
                      {" \u2014 "}
                      {formatSemester(e.sem)}
                      {"\uC5D0 \uC2E0\uCCAD, \uD3B8\uC81C\uD45C\uC5D0\uB294 "}
                      {e.offered.map(formatSemester).join(" · ")}
                      {" \uAC1C\uC124 ("}
                      {e.count}명)
                    </div>
                  ))}
                  {a.notInCurriculum.map((e, i) => (
                    <div key={"n" + i}>
                      {"\u00B7 "}
                      {e.name}
                      {" ("}
                      {formatSemester(e.sem)}) — 편제표에 없는 과목 ({e.count}명)
                    </div>
                  ))}
                  <div
                    style={{
                      color: MUTED,
                      marginTop: 2,
                    }}
                  >
                    올린 편제표가 이 학년도의 최신 편제표인지, 파일의 학기가 맞는지 확인해주세요. 학기가 다른 과목은
                    신청한 학기 기준으로 학점만 반영되고 택N 계산에서는 빠집니다.
                  </div>
                </div>
              </details>
            ))}
          {!a && !hasCurriculum && (
            <div
              style={{
                color: MUTED,
              }}
            >
              편제표를 올리면 이 파일의 과목이 편제표와 맞는지 여기서 바로 확인됩니다.
            </div>
          )}
        </div>
      )}
    </div>
  );
}
function RegSourceZone({
  source,
  regFiles,
  regFileInfos,
  onFiles,
  onRemove,
  onRelabel,
  onSemester,
  onMove,
  hasCurriculum,
  onToggleDisabled,
}) {
  const meta = REG_SOURCES[source];
  const mine = regFiles.filter((f) => f.source === source);
  return (
    <div
      style={{
        background: "#fff",
        border: `1px solid ${LINE}`,
        borderTop: `3px solid ${meta.color}`,
        borderRadius: 10,
        padding: 16,
        display: "flex",
        flexDirection: "column",
        minWidth: 0,
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          marginBottom: 6,
        }}
      >
        <span
          style={{
            fontSize: 12,
            fontWeight: 700,
            color: "#fff",
            background: meta.color,
            borderRadius: 5,
            padding: "2px 8px",
          }}
        >
          {meta.badge}
        </span>
        <span
          style={{
            fontSize: 14.5,
            fontWeight: 700,
          }}
        >
          {meta.title}
        </span>
      </div>
      <div
        style={{
          fontSize: 12.3,
          color: MUTED,
          lineHeight: 1.6,
          marginBottom: 10,
        }}
      >
        {meta.desc}
        <div
          style={{
            marginTop: 4,
          }}
        >
          <b
            style={{
              color: INK,
            }}
          >
            알아보는 법
          </b>
          {" \u00B7 "}
          {meta.signature}
        </div>
      </div>
      <RegFileDropzone onFiles={(fs) => onFiles(fs, source)} label={meta.dropLabel} />
      {mine.length > 0 && (
        <div
          style={{
            marginTop: 12,
            display: "flex",
            flexDirection: "column",
            gap: 8,
          }}
        >
          {mine.map((f) => (
            <RegFileRow
              key={f.id}
              f={f}
              info={regFileInfos.get(f.id)}
              onRemove={onRemove}
              onRelabel={onRelabel}
              onSemester={onSemester}
              onMove={onMove}
              hasCurriculum={hasCurriculum}
              onToggleDisabled={onToggleDisabled}
            />
          ))}
        </div>
      )}
    </div>
  );
}
function RosterPage({
  regFiles,
  regFileInfos,
  handleRegFiles,
  removeRegFile,
  relabelRegFile,
  setRegFileSemester,
  moveRegFile,
  semesterConflicts,
  setSemesterSource,
  rawRecords,
  ambiguousNames,
  identityKeyFor,
  mergeGroups,
  checked,
  toggleCheck,
  mergeChecked,
  unmerge,
  unresolvedCount,
  studentCount,
  onNext,
  canProceed,
  semesterSummary,
  hasCurriculum,
  classMappingFiles,
  handleClassMappingFiles,
  removeClassMappingFile,
  mappedResolvedCount,
  hakjeomLinkedCount,
  rosterMode,
  setRosterMode,
  surveyPanel,
  surveyFileCount,
  sourcePanel,
  sourceSemesters,
  onToggleRegFileDisabled,
}) {
  const byName = useMemo(() => {
    const m = new Map();
    Array.from(ambiguousNames).forEach((name) => {
      m.set(
        name,
        rawRecords.filter((r) => r.name === name),
      );
    });
    return m;
  }, [ambiguousNames, rawRecords]);
  const mergedUids = useMemo(() => new Set(mergeGroups.flatMap((g) => g.recordIds)), [mergeGroups]);
  const zoneProps = {
    regFiles,
    regFileInfos,
    onFiles: handleRegFiles,
    onRemove: removeRegFile,
    onRelabel: relabelRegFile,
    onSemester: setRegFileSemester,
    onMove: moveRegFile,
    hasCurriculum,
    onToggleDisabled: onToggleRegFileDisabled,
  };
  return (
    <div>
      {!hasCurriculum && (
        <div
          style={{
            display: "flex",
            gap: 10,
            alignItems: "flex-start",
            background: "#fff",
            border: `1px solid ${LINE}`,
            borderRadius: 10,
            padding: "14px 16px",
            fontSize: 13,
            color: MUTED,
            marginBottom: 16,
          }}
        >
          <Info
            size={16}
            style={{
              flexShrink: 0,
              marginTop: 1,
            }}
          />
          편제표를 먼저 올리면 학기마다 어떤 자료가 필요한지(학기별 자료 출처)와, 파일마다 편제표와 어긋나는 과목을
          여기서 바로 볼 수 있어요.
        </div>
      )}
      {hasCurriculum && sourcePanel}
      <RosterModePicker
        mode={rosterMode}
        setMode={setRosterMode}
        uploadCount={regFiles.length}
        surveyCount={surveyFileCount}
        uploadSems={sourceSemesters.upload}
        surveySems={sourceSemesters.survey}
      />
      {rosterMode === "survey" ? (
        surveyPanel
      ) : (
        <>
          <div
            style={{
              marginBottom: 16,
              fontSize: 13.5,
              color: MUTED,
              lineHeight: 1.6,
            }}
          >
            {"\uC218\uAC15\uC2E0\uCCAD \uACB0\uACFC \uD30C\uC77C\uC740 "}
            <b
              style={{
                color: INK,
              }}
            >
              어디서 내려받았는지에 따라 양식이 달라서
            </b>
            {
              " \uCE78\uC744 \uB098\uB220 \uBC1B\uC2B5\uB2C8\uB2E4. \uC544\uB798 \uB450 \uCE78 \uC911 \uB9DE\uB294 \uACF3\uC5D0 \uC62C\uB824\uC8FC\uC138\uC694. \uB2E4\uB978 \uC591\uC2DD\uC758 \uD30C\uC77C\uC744 \uC62C\uB9AC\uBA74 \uACC4\uC0B0\uC5D0 \uB123\uC9C0 \uC54A\uACE0 \uC54C\uB824\uB4DC\uB9BD\uB2C8\uB2E4. \uB204\uB9AC\uC9D1 \uD30C\uC77C, \uC555\uD540 \uD30C\uC77C, \uD559\uC0DD \uAE30\uCD08\uC870\uC0AC\uB97C "
            }
            <b
              style={{
                color: INK,
              }}
            >
              학기마다 섞어 써도 됩니다
            </b>
            (예: 2학년 과목은 기초조사, 3학년 과목은 고교학점제 누리집). 어느 학기에 어느 자료를 쓰는지는 위 ‘학기별
            자료 출처’에서 확인합니다.
          </div>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(340px, 1fr))",
              gap: 14,
              marginBottom: 20,
            }}
          >
            <RegSourceZone source="hakjeom" {...zoneProps} />
            <RegSourceZone source="apin" {...zoneProps} />
          </div>
        </>
      )}
      {rawRecords.length > 0 && (
        <>
          <div
            style={{
              display: "flex",
              gap: 12,
              marginTop: 4,
              marginBottom: 20,
              flexWrap: "wrap",
            }}
          >
            <StatCard label="원본 신청 레코드" value={rawRecords.length} />
            <StatCard label="확인된 학생 수" value={studentCount} tone={ACCENT} />
            <StatCard label="동명이인 확인 필요" value={unresolvedCount} tone={unresolvedCount > 0 ? WARN : OK} />
          </div>
          {hakjeomLinkedCount > 0 && (
            <div
              style={{
                marginTop: -8,
                marginBottom: 16,
                fontSize: 12.5,
                color: OK,
                fontWeight: 600,
                lineHeight: 1.55,
              }}
            >
              {
                "\u2713 \uACE0\uAD50\uD559\uC810\uC81C \uB204\uB9AC\uC9D1 \uD30C\uC77C\uC5D0 1\uD559\uB144 \uD559\uBC88\uB3C4 \uD568\uAED8 \uC801\uD600 \uC788\uC5B4, \uB3D9\uBA85\uC774\uC778 "
              }
              {hakjeomLinkedCount}명을 1학년 기록과 자동으로 연결했습니다.
            </div>
          )}
          {ambiguousNames.size > 0 && (
            <div
              style={{
                background: "#fff",
                border: `1px solid ${LINE}`,
                borderRadius: 10,
                padding: 16,
                marginBottom: 20,
              }}
            >
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 700,
                  marginBottom: 4,
                }}
              >
                학급 편성 변경 내역으로 자동 연결
              </div>
              <div
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  marginBottom: 6,
                  lineHeight: 1.5,
                }}
              >
                진급 시 교무부에서 배포하는 반 편성 자료를 올리면, 동명이인 학생을 학번이 달라도 자동으로 연결합니다. 열
                순서는{" "}
                <b
                  style={{
                    color: INK,
                  }}
                >
                  1학년 반(숫자) · 1학년 번호(숫자) · 이름 · 2학년 반(숫자) · 2학년 번호(숫자)
                </b>
                {
                  ' \uC21C\uC73C\uB85C \uB9CC\uB4E4\uC5B4\uC8FC\uC138\uC694 (\uD5E4\uB354 \uD14D\uC2A4\uD2B8\uB294 \uC608\uC2DC\uCC98\uB7FC "N\uD559\uB144 \uBC18"/"N\uD559\uB144 \uBC88\uD638"\uAC00 \uB4E4\uC5B4\uAC00\uBA74 \uC21C\uC11C\uAC00 \uB2EC\uB77C\uB3C4 \uC778\uC2DD\uB429\uB2C8\uB2E4). \uB610\uB294 "10305"(1\uD559\uB144 3\uBC18 5\uBC88)\uCC98\uB7FC \uD55C \uCE78\uC5D0 \uD569\uCCD0\uC9C4 5\uC790\uB9AC \uD559\uBC88 \uCF54\uB4DC\uB97C \uBCC0\uACBD \uC804/\uD6C4 \uB450 \uC5F4\uC5D0 \uB123\uC5B4\uB3C4 \uC778\uC2DD\uB429\uB2C8\uB2E4.'
                }
              </div>
              <button
                onClick={downloadClassMappingSample}
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 5,
                  fontSize: 12,
                  fontWeight: 700,
                  color: ACCENT,
                  background: "none",
                  border: "none",
                  padding: "2px 0 10px",
                  cursor: "pointer",
                }}
              >
                <Download size={12} />
                {" \uC0D8\uD50C \uD30C\uC77C \uB2E4\uC6B4\uB85C\uB4DC"}
              </button>
              <RegFileDropzone onFiles={handleClassMappingFiles} label="끌어다 놓거나" />
              {classMappingFiles.length > 0 && (
                <div
                  style={{
                    marginTop: 10,
                    display: "flex",
                    flexDirection: "column",
                    gap: 6,
                  }}
                >
                  {classMappingFiles.map((f, idx) => (
                    <div
                      key={idx}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        fontSize: 12.5,
                        background: PAPER,
                        borderRadius: 6,
                        padding: "6px 10px",
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 8,
                          minWidth: 0,
                        }}
                      >
                        {f.status === "error" ? (
                          <AlertCircle
                            size={14}
                            color={WARN}
                            style={{
                              flexShrink: 0,
                            }}
                          />
                        ) : (
                          <CheckCircle2
                            size={14}
                            color={OK}
                            style={{
                              flexShrink: 0,
                            }}
                          />
                        )}
                        <span
                          style={{
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {f.name}
                        </span>
                        {f.status === "ok" && (
                          <span
                            style={{
                              color: MUTED,
                              flexShrink: 0,
                            }}
                          >
                            {"\u00B7 "}
                            {f.count}건
                          </span>
                        )}
                        {f.status === "error" && (
                          <span
                            style={{
                              color: WARN,
                              flexShrink: 0,
                            }}
                          >
                            {"\u00B7 "}
                            {f.error}
                          </span>
                        )}
                      </div>
                      <button
                        onClick={() => removeClassMappingFile(idx)}
                        style={{
                          border: "none",
                          background: "none",
                          cursor: "pointer",
                          color: MUTED,
                          flexShrink: 0,
                        }}
                      >
                        <X size={14} />
                      </button>
                    </div>
                  ))}
                </div>
              )}
              {mappedResolvedCount > 0 && (
                <div
                  style={{
                    marginTop: 10,
                    fontSize: 12.5,
                    color: OK,
                    fontWeight: 600,
                  }}
                >
                  {"\u2713 \uB3D9\uBA85\uC774\uC778 "}
                  {mappedResolvedCount}명이 이 자료로 자동 연결되었습니다.
                </div>
              )}
            </div>
          )}
          {ambiguousNames.size === 0 ? (
            <div
              style={{
                display: "flex",
                gap: 10,
                alignItems: "center",
                background: OK_BG,
                borderRadius: 10,
                padding: "14px 16px",
                fontSize: 13,
                color: OK,
              }}
            >
              <CheckCircle2 size={16} />
              같은 파일 안에서 이름이 겹치는 학생이 없습니다. 모든 기록이 이름 기준으로 자동 연결되었습니다.
            </div>
          ) : (
            <div>
              <div
                style={{
                  display: "flex",
                  gap: 8,
                  alignItems: "flex-start",
                  marginBottom: 14,
                  fontSize: 13,
                  color: MUTED,
                  lineHeight: 1.55,
                }}
              >
                <Info
                  size={15}
                  style={{
                    flexShrink: 0,
                    marginTop: 1,
                  }}
                />
                <span>
                  {
                    "\uC544\uB798 \uC774\uB984\uB4E4\uC740 \uD55C \uD30C\uC77C \uC548\uC5D0 \uAC19\uC740 \uC774\uB984\uC774 2\uBA85 \uC774\uC0C1 \uC788\uC2B5\uB2C8\uB2E4. \uD559\uB144\u00B7\uBC18\u00B7\uBC88\uD638\uAC00 \uAC19\uC740 \uAE30\uB85D\uB07C\uB9AC\uB294 \uC790\uB3D9\uC73C\uB85C \uBB36\uC5C8\uC9C0\uB9CC, \uD559\uB144\uC774 \uBC14\uB00C\uBA74 \uBC18\u00B7\uBC88\uD638\uAC00 \uB2EC\uB77C\uC838\uC11C(\uC608: \uC555\uD540 "
                  }
                  <b
                    style={{
                      color: INK,
                    }}
                  >
                    1학년 3반 5번
                  </b>
                  {" \u2194 \uACE0\uAD50\uD559\uC810\uC81C "}
                  <b
                    style={{
                      color: INK,
                    }}
                  >
                    학번 20305
                  </b>
                  {
                    ') \uADF8 \uC0AC\uC774\uB294 \uC790\uB3D9\uC73C\uB85C \uC774\uC744 \uC218 \uC5C6\uC2B5\uB2C8\uB2E4. \uAC19\uC740 \uD559\uC0DD\uC778 \uC904\uB07C\uB9AC \uCCB4\uD06C\uD55C \uB4A4 "\uC120\uD0DD \uD56D\uBAA9 \uBCD1\uD569"\uC744 \uB20C\uB7EC\uC8FC\uC138\uC694.'
                  }
                </span>
              </div>
              <div
                style={{
                  display: "flex",
                  flexDirection: "column",
                  gap: 14,
                }}
              >
                {Array.from(byName.entries()).map(([name, recs]) => {
                  const groupsForName = mergeGroups.filter((g) =>
                    g.recordIds.some((uid) => recs.some((r) => r.uid === uid)),
                  );
                  const units = new Map();
                  recs.forEach((r) => {
                    if (mergedUids.has(r.uid)) return;
                    const k = identityKeyFor(r);
                    if (!units.has(k)) units.set(k, []);
                    units.get(k).push(r);
                  });
                  const unitList = Array.from(units.entries()).map(([key, rs]) => ({
                    key,
                    recs: rs,
                    mapped: key.startsWith("mapped:"),
                  }));
                  const checkedUnits = unitList.filter((u) => checked[u.key]);
                  return (
                    <div
                      key={name}
                      style={{
                        background: "#fff",
                        border: `1px solid ${LINE}`,
                        borderRadius: 10,
                        padding: 16,
                      }}
                    >
                      <div
                        style={{
                          display: "flex",
                          alignItems: "center",
                          justifyContent: "space-between",
                          marginBottom: 10,
                          gap: 8,
                          flexWrap: "wrap",
                        }}
                      >
                        <div
                          style={{
                            display: "flex",
                            alignItems: "center",
                            gap: 8,
                            fontWeight: 700,
                            fontSize: 14,
                          }}
                        >
                          <Users size={15} color={WARN} />
                          {name}
                          <span
                            style={{
                              fontWeight: 400,
                              fontSize: 12,
                              color: MUTED,
                            }}
                          >
                            {"\u00B7 \uCD1D "}
                            {recs.length}건 발견
                          </span>
                        </div>
                        {checkedUnits.length > 0 && (
                          <button
                            disabled={checkedUnits.length < 2}
                            onClick={() =>
                              mergeChecked(
                                checkedUnits.flatMap((u) => u.recs.map((r) => r.uid)),
                                checkedUnits.map((u) => u.key),
                              )
                            }
                            style={{
                              display: "flex",
                              alignItems: "center",
                              gap: 5,
                              fontSize: 12.5,
                              fontWeight: 700,
                              color: "#fff",
                              background: checkedUnits.length < 2 ? "#B8B4A9" : ACCENT,
                              border: "none",
                              borderRadius: 7,
                              padding: "6px 12px",
                              cursor: checkedUnits.length < 2 ? "default" : "pointer",
                            }}
                          >
                            <Check size={13} />{" "}
                            {checkedUnits.length < 2 ? "2줄 이상 선택하세요" : "선택 항목 병합 (같은 학생)"}
                          </button>
                        )}
                      </div>
                      {groupsForName.map((g) => (
                        <div
                          key={g.id}
                          style={{
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "space-between",
                            background: OK_BG,
                            borderRadius: 7,
                            padding: "8px 12px",
                            marginBottom: 6,
                            fontSize: 12.5,
                            gap: 8,
                          }}
                        >
                          <span
                            style={{
                              color: OK,
                              lineHeight: 1.5,
                            }}
                          >
                            병합됨:{" "}
                            {g.recordIds
                              .map((uid) => recs.find((r) => r.uid === uid))
                              .filter(Boolean)
                              .map((r) => {
                                var _a, _b;
                                return `${r.source === SURVEY_SOURCE_ID ? SURVEY_SOURCE_META.short : (_b = (_a = REG_SOURCES[r.source]) === null || _a === void 0 ? void 0 : _a.short) !== null && _b !== void 0 ? _b : ""} ${recordSemesterLabel(r)} ${r.idText}`;
                              })
                              .join(" + ")}
                          </span>
                          <button
                            onClick={() => unmerge(g.id)}
                            style={{
                              border: "none",
                              background: "none",
                              color: OK,
                              cursor: "pointer",
                              fontSize: 12,
                              flexShrink: 0,
                            }}
                          >
                            병합 해제
                          </button>
                        </div>
                      ))}
                      {unitList.map((u) => (
                        <label
                          key={u.key}
                          style={{
                            display: "flex",
                            alignItems: "flex-start",
                            gap: 10,
                            padding: "8px 10px",
                            borderRadius: 7,
                            cursor: "pointer",
                            background: checked[u.key] ? ACCENT_BG : PAPER,
                            marginBottom: 6,
                            fontSize: 12.5,
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={!!checked[u.key]}
                            onChange={() => toggleCheck(u.key)}
                            style={{
                              marginTop: 2,
                            }}
                          />
                          <div
                            style={{
                              display: "flex",
                              flexDirection: "column",
                              gap: 3,
                              minWidth: 0,
                            }}
                          >
                            {u.recs.map((r) => (
                              <div
                                key={r.uid}
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: 8,
                                  flexWrap: "wrap",
                                }}
                              >
                                <SourceTag source={r.source} />
                                <span
                                  style={{
                                    fontWeight: 600,
                                  }}
                                >
                                  {recordSemesterLabel(r)}
                                </span>
                                <span
                                  style={{
                                    color: MUTED,
                                  }}
                                >
                                  {r.idText}
                                </span>
                                <span
                                  style={{
                                    color: MUTED,
                                  }}
                                >
                                  {"\u00B7 \uC2E0\uCCAD "}
                                  {r.subjects.length}과목
                                </span>
                              </div>
                            ))}
                            {u.recs.length > 1 && (
                              <div
                                style={{
                                  fontSize: 11.5,
                                  color: ACCENT,
                                }}
                              >
                                {u.mapped ? "학급 편성 자료로 자동 연결됨" : "학년·반·번호가 같아 자동으로 묶임"}
                              </div>
                            )}
                          </div>
                        </label>
                      ))}
                    </div>
                  );
                })}
              </div>
            </div>
          )}
        </>
      )}
      <div
        style={{
          display: "flex",
          justifyContent: "flex-end",
          marginTop: 24,
        }}
      >
        <button
          disabled={!canProceed}
          onClick={onNext}
          style={{
            display: "flex",
            alignItems: "center",
            gap: 6,
            fontSize: 13.5,
            fontWeight: 700,
            color: "#fff",
            background: !canProceed ? "#B8B4A9" : ACCENT,
            border: "none",
            borderRadius: 8,
            padding: "10px 18px",
            cursor: !canProceed ? "default" : "pointer",
          }}
        >
          {"\uB2E4\uC74C: \uC878\uC5C5\uC694\uAC74 \uC810\uAC80 "}
          <ArrowRight size={15} />
        </button>
      </div>
      {!canProceed && (
        <div
          style={{
            textAlign: "right",
            fontSize: 12,
            color: WARN,
            marginTop: 6,
          }}
        >
          편제표를 먼저 업로드해주세요.
        </div>
      )}
    </div>
  );
}
// ================= PAGE 2 · 학생 기초조사 =================
// <script src="주소&callback=이름">으로 JSON을 받습니다. fetch가 다른 사이트 요청으로 막힐 때의 예비 방식 (파일로 연 화면에서도 동작)
function loadJsonp(url, timeoutMs = 25000) {
  return new Promise((resolve, reject) => {
    const name = "__survey_cb_" + Math.random().toString(36).slice(2);
    const script = document.createElement("script");
    const done = (fn, v) => {
      clearTimeout(timer);
      delete window[name];
      script.remove();
      fn(v);
    };
    const timer = setTimeout(() => done(reject, new Error("JSONP timeout")), timeoutMs);
    window[name] = (data) => done(resolve, data);
    script.onerror = () => done(reject, new TypeError("JSONP failed"));
    script.src = url + (url.includes("?") ? "&" : "?") + "callback=" + name;
    document.body.appendChild(script);
  });
}
async function copyTextToClipboard(text) {
  try {
    if (navigator.clipboard && window.isSecureContext) {
      await navigator.clipboard.writeText(text);
      return true;
    }
  } catch (e) {
    // 파일로 연 화면 등에서 막히면 아래 방식으로 다시 시도
  }
  const ta = document.createElement("textarea");
  ta.value = text;
  ta.setAttribute("readonly", "");
  ta.style.position = "fixed";
  ta.style.top = "-2000px";
  document.body.appendChild(ta);
  ta.select();
  let ok = false;
  try {
    ok = document.execCommand("copy");
  } catch (e) {
    ok = false;
  }
  document.body.removeChild(ta);
  return ok;
}
function SurveyStep({ no, title, desc, children, tone }) {
  return (
    <div
      style={{
        ...CARD,
        padding: 16,
        marginBottom: 14,
        borderTop: `3px solid ${tone || "#B7791F"}`,
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          marginBottom: desc ? 4 : 10,
        }}
      >
        <span
          style={{
            width: 22,
            height: 22,
            borderRadius: 999,
            background: tone || "#B7791F",
            color: "#fff",
            fontSize: 12.5,
            fontWeight: 800,
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            flexShrink: 0,
          }}
        >
          {no}
        </span>
        <span
          style={{
            fontSize: 15,
            fontWeight: 800,
          }}
        >
          {title}
        </span>
      </div>
      {desc && (
        <div
          style={{
            fontSize: 12.5,
            color: MUTED,
            lineHeight: 1.6,
            marginBottom: 10,
          }}
        >
          {desc}
        </div>
      )}
      {children}
    </div>
  );
}
function SurveyScopeStep({ config, setConfig, poolDefs, payload, admissionYear }) {
  const available = surveySemestersWithGroups(poolDefs);
  const isOn = (sem) => (config.semesters ? config.semesters.includes(sem) : available.includes(sem));
  const toggle = (sem) =>
    setConfig((c) => {
      const cur = c.semesters ? c.semesters : available;
      return {
        ...c,
        semesters: cur.includes(sem) ? cur.filter((s) => s !== sem) : [...cur, sem].sort(compareSemester),
      };
    });
  const groupCount = payload.semesters.reduce((a, s) => a + s.groups.length, 0);
  const optionCount = payload.semesters.reduce((a, s) => a + s.groups.reduce((b, g) => b + g.options.length, 0), 0);
  const sel = (missing) => ({
    ...inputStyle,
    fontSize: 13,
    padding: "5px 8px",
    borderColor: missing ? WARN : LINE,
    background: missing ? WARN_BG : "#fff",
  });
  return (
    <SurveyStep
      no={1}
      title="조사 범위"
      desc={
        <>
          {
            "\uD559\uC0DD\uC774 \uC9C1\uC811 \uCC44\uC6B8 \uD559\uAE30\uB97C \uACE0\uB985\uB2C8\uB2E4. \uD3B8\uC81C\uD45C\uC5D0 "
          }
          <b
            style={{
              color: INK,
            }}
          >
            학생선택 과목이 있는 학기
          </b>
          만 고를 수 있고, 학교지정 과목은 묻지 않습니다.
        </>
      }
    >
      <div
        style={{
          display: "flex",
          gap: 14,
          flexWrap: "wrap",
          marginBottom: 12,
        }}
      >
        {SEMESTER_CODES.map((sem) => {
          const enabled = available.includes(sem);
          return (
            <label
              key={sem}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 6,
                fontSize: 14,
                color: enabled ? INK : "#B8B4A9",
                cursor: enabled ? "pointer" : "default",
              }}
            >
              <input
                type="checkbox"
                disabled={!enabled}
                checked={enabled && isOn(sem)}
                onChange={() => toggle(sem)}
                style={{
                  width: 16,
                  height: 16,
                }}
              />
              {formatSemester(sem)}
              {!enabled && (
                <span
                  style={{
                    fontSize: 12,
                  }}
                >
                  (편제표에 학생선택 과목 없음)
                </span>
              )}
            </label>
          );
        })}
      </div>
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 6,
          flexWrap: "wrap",
          fontSize: 13,
          marginBottom: 6,
        }}
      >
        <span
          style={{
            color: config.grade && config.term ? MUTED : WARN,
            fontWeight: config.grade && config.term ? 400 : 700,
          }}
        >
          조사하는 때(지금):
        </span>
        <select
          value={config.grade || ""}
          onChange={(e) =>
            setConfig((c) => ({
              ...c,
              grade: Number(e.target.value) || null,
            }))
          }
          style={sel(!config.grade)}
        >
          <option value="">학년 선택</option>
          {[1, 2, 3].map((g) => (
            <option key={g} value={g}>
              {g}학년
            </option>
          ))}
        </select>
        <select
          value={config.term || ""}
          onChange={(e) =>
            setConfig((c) => ({
              ...c,
              term: Number(e.target.value) || null,
            }))
          }
          style={sel(!config.term)}
        >
          <option value="">학기 선택</option>
          {[1, 2].map((t) => (
            <option key={t} value={t}>
              {t}학기
            </option>
          ))}
        </select>
        <span
          style={{
            color: MUTED,
            fontSize: 12.3,
          }}
        >
          → 학생 화면에서 지난 학기는 ‘들은 과목’, 이번 학기는 ‘지금 듣는 과목’, 다음 학기는 ‘신청한 과목’으로
          표시됩니다.{" "}
          {config.grade ? (
            <>
              {"\uBC18\u00B7\uBC88\uD638\uB294 "}
              <b
                style={{
                  color: INK,
                }}
              >
                {config.grade}학년 기준
              </b>
              으로 받습니다.
            </>
          ) : (
            <b
              style={{
                color: WARN,
              }}
            >
              학생들의 지금 학년을 고르세요 (반·번호를 이 학년 기준으로 받습니다).
            </b>
          )}
        </span>
      </div>
      <div
        style={{
          fontSize: 13,
          color: ACCENT,
          fontWeight: 700,
          marginTop: 8,
        }}
      >
        {"\u2192 \uD559\uAE30 "}
        {payload.semesters.length}
        {"\uAC1C \u00B7 \uC120\uD0DD \uBB36\uC74C "}
        {groupCount}
        {"\uAC1C \u00B7 \uACE0\uB97C \uC218 \uC788\uB294 \uACFC\uBAA9 "}
        {optionCount}개
      </div>
    </SurveyStep>
  );
}
function SurveyRosterStep({ config, setConfig }) {
  const [fillSize, setFillSize] = useState("");
  const count = config.classes.length;
  const setCount = (n) =>
    setConfig((c) => {
      const want = Math.max(0, Math.min(30, Math.floor(Number(n)) || 0));
      const byC = new Map(c.classes.map((k) => [k.c, k]));
      return {
        ...c,
        classes: Array.from(
          {
            length: want,
          },
          (_, i) => ({
            c: i + 1,
            size: (byC.get(i + 1) || {}).size || 0,
            skip: (byC.get(i + 1) || {}).skip || [],
          }),
        ),
      };
    });
  const setSize = (cNum, size) =>
    setConfig((c) => ({
      ...c,
      classes: c.classes.map((k) =>
        k.c === cNum
          ? {
              ...k,
              size: Math.max(0, Math.min(60, Math.floor(Number(size)) || 0)),
            }
          : k,
      ),
    }));
  const total = config.classes.reduce(
    (a, k) => a + (k.size || 0) - (k.skip || []).filter((n) => n <= (k.size || 0)).length,
    0,
  );
  const setSkip = (cNum, list) =>
    setConfig((c) => ({
      ...c,
      classes: c.classes.map((k) =>
        k.c === cNum
          ? {
              ...k,
              skip: list,
            }
          : k,
      ),
    }));
  return (
    <SurveyStep
      no={2}
      title="학생 명단 (반별 인원)"
      desc="학생 화면의 반·번호 버튼과, 아직 내지 않은 번호를 찾는 데 씁니다. 이름은 받지 않습니다. 인원은 마지막 번호로 적고, 전학·자퇴 등으로 비어 있는 번호는 ‘결번’ 칸에 적으세요(예: 7, 15). 결번은 학생 화면에서 누를 수 없고 안 낸 학생에서도 빠집니다."
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
          flexWrap: "wrap",
          fontSize: 13,
          marginBottom: 12,
        }}
      >
        <span>반 수</span>
        <input
          type="number"
          min={0}
          max={30}
          value={count || ""}
          onChange={(e) => setCount(e.target.value)}
          placeholder="예: 8"
          style={{
            ...inputStyle,
            width: 70,
            fontSize: 13,
          }}
        />
        <span
          style={{
            color: LINE,
          }}
        >
          |
        </span>
        <span>모든 반</span>
        <input
          type="number"
          min={1}
          max={60}
          value={fillSize}
          onChange={(e) => setFillSize(e.target.value)}
          placeholder="28"
          style={{
            ...inputStyle,
            width: 64,
            fontSize: 13,
          }}
        />
        <span>명으로</span>
        <button
          disabled={!count || !(Number(fillSize) >= 1)}
          onClick={() =>
            setConfig((c) => ({
              ...c,
              classes: c.classes.map((k) => ({
                ...k,
                size: Math.min(60, Math.floor(Number(fillSize))),
              })),
            }))
          }
          style={buttonStyle(!count || !(Number(fillSize) >= 1) ? "disabled" : "soft")}
        >
          한꺼번에 채우기
        </button>
        {count > 0 && (
          <span
            style={{
              color: MUTED,
            }}
          >
            {"\u00B7 \uCD1D "}
            {total}명
          </span>
        )}
      </div>
      {count > 0 && (
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "repeat(auto-fill, minmax(260px, 1fr))",
            gap: 8,
          }}
        >
          {config.classes.map((k) => (
            <div
              key={k.c}
              style={{
                display: "grid",
                gridTemplateColumns: "34px 58px 16px 30px minmax(0, 1fr)",
                alignItems: "center",
                gap: 5,
                fontSize: 13,
                background: PAPER,
                borderRadius: 7,
                padding: "6px 9px",
              }}
            >
              <b>{k.c}반</b>
              <input
                type="number"
                min={1}
                max={60}
                value={k.size || ""}
                onChange={(e) => setSize(k.c, e.target.value)}
                style={{
                  ...inputStyle,
                  width: 58,
                  fontSize: 13,
                  borderColor: k.size >= 1 ? LINE : WARN,
                }}
              />
              <span>명</span>
              <span
                style={{
                  fontSize: 12,
                  color: MUTED,
                }}
              >
                결번
              </span>
              <SkipInput value={k.skip} max={k.size} onChange={(list) => setSkip(k.c, list)} />
            </div>
          ))}
        </div>
      )}
    </SurveyStep>
  );
}
function SurveyTextStep({ config, setConfig, onPreview, canPreview }) {
  const field = (label, key, placeholder, max) => (
    <label
      style={{
        display: "flex",
        flexDirection: "column",
        gap: 4,
        fontSize: 12.5,
        color: MUTED,
        flex: 1,
        minWidth: 220,
      }}
    >
      {label}
      <input
        value={config[key]}
        maxLength={max}
        onChange={(e) =>
          setConfig((c) => ({
            ...c,
            [key]: e.target.value,
          }))
        }
        placeholder={placeholder}
        style={{
          ...inputStyle,
          fontSize: 13.5,
          padding: "8px 10px",
        }}
      />
    </label>
  );
  return (
    <SurveyStep
      no={3}
      title="학생 화면에 보일 안내"
      desc="학생 휴대폰 화면 맨 위에 나옵니다. 비워 두면 기본 제목(선택과목 기초조사)이 쓰입니다."
    >
      <div
        style={{
          display: "flex",
          gap: 10,
          flexWrap: "wrap",
          marginBottom: 10,
        }}
      >
        {field("조사 제목", "title", "예: 2025학년도 입학생 선택과목 기초조사", 60)}
        {field("마감 안내", "deadline", "예: 9월 20일(금) 오후 4시까지", 60)}
      </div>
      <label
        style={{
          display: "flex",
          flexDirection: "column",
          gap: 4,
          fontSize: 12.5,
          color: MUTED,
          marginBottom: 12,
        }}
      >
        안내 문구 (선택)
        <textarea
          value={config.notice}
          maxLength={400}
          rows={2}
          onChange={(e) =>
            setConfig((c) => ({
              ...c,
              notice: e.target.value,
            }))
          }
          placeholder="예: 헷갈리면 학생부나 수강신청 확인서를 보고 고르세요. 모르는 것은 담임 선생님께 물어보세요."
          style={{
            ...inputStyle,
            fontSize: 13.5,
            padding: "8px 10px",
            resize: "vertical",
          }}
        />
      </label>
      <button
        disabled={!canPreview}
        onClick={onPreview}
        style={buttonStyle(canPreview ? "soft" : "disabled", {
          padding: "8px 14px",
        })}
      >
        📱 학생 화면 미리보기
      </button>
      {!canPreview && (
        <span
          style={{
            fontSize: 12,
            color: MUTED,
            marginLeft: 8,
          }}
        >
          편제표와 반별 인원을 넣으면 미리 볼 수 있습니다.
        </span>
      )}
    </SurveyStep>
  );
}
function SurveyPreviewModal({ html, onClose }) {
  return (
    <div
      onClick={onClose}
      style={{
        position: "fixed",
        inset: 0,
        background: "rgba(28,35,51,0.45)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 70,
        padding: 12,
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: "#fff",
          borderRadius: 16,
          padding: 14,
          fontFamily: FONT,
          boxShadow: "0 12px 32px rgba(0,0,0,0.25)",
          maxHeight: "96vh",
          display: "flex",
          flexDirection: "column",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 10,
            marginBottom: 8,
          }}
        >
          <div>
            <div
              style={{
                fontSize: 14,
                fontWeight: 800,
              }}
            >
              학생 화면 미리보기
            </div>
            <div
              style={{
                fontSize: 11.5,
                color: MUTED,
              }}
            >
              학생 휴대폰에 보이는 모습입니다. 눌러 봐도 제출되지 않습니다.
            </div>
          </div>
          <button onClick={onClose} style={buttonStyle("ghost")}>
            닫기
          </button>
        </div>
        <iframe
          title="학생 화면 미리보기"
          srcDoc={html}
          sandbox="allow-scripts"
          style={{
            width: 390,
            maxWidth: "88vw",
            height: 740,
            maxHeight: "calc(96vh - 70px)",
            border: `1px solid ${LINE}`,
            borderRadius: 12,
            background: PAPER,
          }}
        />
      </div>
    </div>
  );
}
function QrSvg({ text, size }) {
  const qr = useMemo(() => {
    try {
      return qrEncodeText(text);
    } catch (e) {
      return null;
    }
  }, [text]);
  if (!qr)
    return (
      <div
        style={{
          width: size,
          height: size,
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          fontSize: 12,
          color: WARN,
          border: `1px solid ${LINE}`,
        }}
      >
        QR 코드를 만들 수 없는 주소
      </div>
    );
  const dim = qr.size + 8;
  return (
    <svg
      width={size}
      height={size}
      viewBox={`0 0 ${dim} ${dim}`}
      shapeRendering="crispEdges"
      style={{
        display: "block",
        background: "#fff",
      }}
    >
      <rect width={dim} height={dim} fill="#fff" />
      <path d={qrSvgPath(qr)} fill="#000" />
    </svg>
  );
}
// 템플릿 시트 주소(…/edit 등)를 '사본 만들기' 주소로 바꿉니다. 비었거나 구글 시트 주소가 아니면 null
function templateCopyUrl(raw) {
  const m = /^https:\/\/docs\.google\.com\/spreadsheets\/d\/([A-Za-z0-9_-]{20,})/.exec(String(raw || "").trim());
  return m ? `https://docs.google.com/spreadsheets/d/${m[1]}/copy` : null;
}
const SURVEY_INSTALL_STEPS = {
  template: ["t1", "t2", "t3"],
  code: ["c1", "c2", "c3", "c4", "c5"],
};
function InstallCheckItem({ no, done, onToggle, title, children }) {
  return (
    <div
      style={{
        display: "flex",
        gap: 10,
        padding: "10px 12px",
        borderRadius: 8,
        border: `1px solid ${done ? OK : LINE}`,
        background: done ? OK_BG : "#fff",
        marginBottom: 8,
      }}
    >
      <label
        style={{ display: "flex", alignItems: "flex-start", paddingTop: 2, cursor: "pointer" }}
        title="다 했으면 체크"
      >
        <input type="checkbox" checked={!!done} onChange={onToggle} style={{ width: 16, height: 16, margin: 0 }} />
      </label>
      <div style={{ flex: 1, minWidth: 0, fontSize: 13, lineHeight: 1.65, color: INK }}>
        <div style={{ fontWeight: 800, marginBottom: 2 }}>
          {no}. {title}
          {done && <span style={{ color: OK, fontWeight: 700, marginLeft: 6, fontSize: 12 }}>✓ 완료</span>}
        </div>
        {children}
      </div>
    </div>
  );
}
// Apps Script 편집기의 [배포 → 새 배포] 창을 실제 배치대로 그린 그림. 고를 곳에 빨간 테두리를 둘렀습니다.
function DeployDialogGuide() {
  const ring = { outline: "3px solid #D93025", outlineOffset: 2, borderRadius: 4 };
  const field = {
    border: "1px solid #DADCE0",
    borderRadius: 4,
    padding: "5px 8px",
    fontSize: 12,
    color: "#202124",
    background: "#fff",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  };
  const cap = { fontSize: 11, color: "#5F6368", margin: "8px 0 3px" };
  return (
    <div
      style={{
        marginTop: 8,
        maxWidth: 520,
        border: `1px solid ${LINE}`,
        borderRadius: 8,
        background: "#fff",
        fontFamily: "Roboto, Arial, sans-serif",
        overflow: "hidden",
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "8px 12px",
          borderBottom: "1px solid #DADCE0",
          fontSize: 12,
          color: "#5F6368",
        }}
      >
        <span>Apps Script 편집기 오른쪽 위</span>
        <span style={ring}>
          <span
            style={{
              display: "inline-block",
              background: "#1A73E8",
              color: "#fff",
              fontWeight: 700,
              borderRadius: 4,
              padding: "4px 10px",
            }}
          >
            배포 ▾
          </span>
        </span>
      </div>
      <div style={{ padding: "4px 12px 0", fontSize: 11.5, color: "#5F6368", textAlign: "right" }}>
        ▸ <b style={{ color: "#202124" }}>새 배포</b>
      </div>
      <div style={{ padding: "8px 14px 12px" }}>
        <div style={{ fontSize: 15, color: "#202124", marginBottom: 6 }}>새 배포</div>
        <div style={{ display: "grid", gridTemplateColumns: "minmax(110px, 38%) 1fr", gap: 12 }}>
          <div style={{ borderRight: "1px solid #DADCE0", paddingRight: 10 }}>
            <div style={{ fontSize: 12, color: "#202124", display: "flex", alignItems: "center", gap: 6 }}>
              유형 선택
              <span style={ring}>
                <span style={{ fontSize: 14, padding: "0 2px" }}>⚙</span>
              </span>
            </div>
            <div style={{ fontSize: 11, color: "#5F6368", margin: "6px 0 3px" }}>톱니바퀴를 누르면</div>
            <span style={{ ...ring, display: "inline-block" }}>
              <span style={{ fontSize: 12, color: "#202124", padding: "2px 6px", display: "inline-block" }}>웹 앱</span>
            </span>
          </div>
          <div>
            <div style={{ ...cap, marginTop: 0 }}>구성 · 설명 (비워 둬도 됨)</div>
            <div style={{ ...field, color: "#9AA0A6" }}>새 설명</div>
            <div style={cap}>다음 사용자 인증 정보로 실행</div>
            <div style={{ ...field, ...ring }}>
              <span>
                <b>나</b> (선생님 이메일)
              </span>
              <span>▾</span>
            </div>
            <div style={cap}>액세스 권한이 있는 사용자</div>
            <div style={{ ...field, ...ring }}>
              <b>모든 사용자</b>
              <span>▾</span>
            </div>
          </div>
        </div>
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8, marginTop: 12, fontSize: 12 }}>
          <span style={{ color: "#1A73E8", padding: "5px 8px" }}>취소</span>
          <span style={ring}>
            <span
              style={{
                display: "inline-block",
                background: "#1A73E8",
                color: "#fff",
                fontWeight: 700,
                borderRadius: 4,
                padding: "5px 14px",
              }}
            >
              배포
            </span>
          </span>
        </div>
      </div>
      <div
        style={{
          borderTop: "1px solid #DADCE0",
          padding: "8px 14px",
          fontSize: 11.5,
          color: "#5F6368",
          lineHeight: 1.6,
        }}
      >
        ‘모든 사용자’를 골라야 학생이 로그인하지 않고 QR 코드로 바로 들어올 수 있습니다. ‘Google 계정이 있는 모든
        사용자’나 ‘나만’을 고르면 학생에게 로그인 화면이 나옵니다. [배포]를 누른 뒤 <b>액세스 승인</b>을 묻는 창이
        나오면 ‘확인하지 않은 앱’ 안내 그림과 같은 순서(고급 → 이동 → 모두 선택 → 계속)로 허용하세요.
      </div>
    </div>
  );
}
// 배포가 끝난 뒤 나오는 '배포가 업데이트되었습니다' 화면. [복사]가 두 개라서 아래쪽(웹 앱 URL)을 짚어 줍니다.
function DeployDonePicture() {
  const ring = { outline: "3px solid #D93025", outlineOffset: 2, borderRadius: 4 };
  const cap = { fontSize: 11, color: "#5F6368", marginTop: 8 };
  const copy = {
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
    color: "#1A73E8",
    fontWeight: 700,
    fontSize: 12,
    padding: "1px 4px",
  };
  return (
    <div
      style={{
        marginTop: 8,
        maxWidth: 520,
        border: `1px solid ${LINE}`,
        borderRadius: 8,
        background: "#fff",
        fontFamily: "Roboto, Arial, sans-serif",
        color: "#3C4043",
        fontSize: 12,
        padding: "12px 14px",
      }}
    >
      <div style={{ fontSize: 15, color: "#202124" }}>새 배포</div>
      <div style={{ marginTop: 6 }}>배포가 업데이트되었습니다.</div>
      <div style={cap}>배포 ID</div>
      <div style={{ color: "#5F6368", fontSize: 11 }}>AKfycby…zeRgA</div>
      <div style={{ marginTop: 2 }}>
        <span style={{ ...copy, color: "#9AA0A6", textDecoration: "line-through" }}>⧉ 복사</span>
        <span style={{ color: "#A2452C", fontSize: 11, marginLeft: 6 }}>← 이건 아닙니다</span>
      </div>
      <div style={{ ...cap, fontSize: 12, color: "#202124", fontWeight: 700 }}>웹 앱</div>
      <div style={{ ...cap, marginTop: 2 }}>URL</div>
      <div style={{ color: "#1A73E8", fontSize: 11, wordBreak: "break-all" }}>
        https://script.google.com/macros/s/AKfycby…/exec
      </div>
      <div style={{ marginTop: 4 }}>
        <span style={{ ...copy, ...ring }}>⧉ 복사</span>
        <span style={{ color: "#A2452C", fontSize: 11, marginLeft: 8 }}>← 이 [복사]를 누르세요</span>
      </div>
      <div style={{ textAlign: "right", marginTop: 10 }}>
        <span
          style={{
            background: "#1A73E8",
            color: "#fff",
            fontWeight: 700,
            borderRadius: 4,
            padding: "5px 14px",
            display: "inline-block",
          }}
        >
          완료
        </span>
      </div>
    </div>
  );
}
// 사본 시트 위쪽 메뉴에 생기는 [📋 기초조사] 메뉴를 연 모습. 첫 번째 항목을 누르라고 짚어 줍니다.
function SheetMenuPicture() {
  const ring = { outline: "3px solid #D93025", outlineOffset: 2, borderRadius: 4 };
  const item = { padding: "7px 14px", fontSize: 12.5, color: "#202124", whiteSpace: "nowrap" };
  return (
    <div
      style={{
        marginTop: 8,
        maxWidth: 520,
        border: `1px solid ${LINE}`,
        borderRadius: 8,
        background: "#fff",
        fontFamily: "Roboto, Arial, sans-serif",
        color: "#3C4043",
        fontSize: 12.5,
        overflow: "hidden",
      }}
    >
      <div
        style={{ display: "flex", gap: 14, padding: "8px 12px", borderBottom: "1px solid #DADCE0", color: "#202124" }}
      >
        <span>파일</span>
        <span>수정</span>
        <span>보기</span>
        <span>삽입</span>
        <span style={{ color: "#9AA0A6" }}>…</span>
        <span>확장 프로그램</span>
        <span>도움말</span>
        <span style={{ ...ring, background: "#E8F0FE", padding: "0 6px", fontWeight: 700 }}>📋 기초조사</span>
      </div>
      <div style={{ display: "flex", justifyContent: "flex-end", padding: "0 12px 10px" }}>
        <div
          style={{
            border: "1px solid #DADCE0",
            borderRadius: 6,
            boxShadow: "0 2px 6px rgba(0,0,0,.15)",
            padding: "4px 0",
            background: "#fff",
          }}
        >
          <div style={{ ...item, ...ring, margin: "4px 8px", fontWeight: 700 }}>설정 붙여넣기 / 바꾸기</div>
          <div style={item}>배포(학생용 주소 만들기) 안내</div>
          <div style={item}>지금 설정 보기</div>
        </div>
      </div>
      <div style={{ padding: "6px 12px 10px", fontSize: 12, color: MUTED, lineHeight: 1.6 }}>
        메뉴는 시트가 열리고 몇 초 뒤 <b>도움말 오른쪽</b>에 생깁니다. 안 보이면 새로고침(F5)하세요. 첫
        화면(시작하기)에도 같은 안내가 적혀 있습니다.
      </div>
    </div>
  );
}
// [배포]를 누른 직후의 '새 배포' 창: 주소 대신 [액세스 승인] 버튼이 먼저 나옵니다. 여기서 멈추는 분이 많아 따로 그립니다.
function DeployAuthPicture() {
  const ring = { outline: "3px solid #D93025", outlineOffset: 2, borderRadius: 4 };
  return (
    <div style={{ marginTop: 10 }}>
      <div style={{ fontSize: 12.5, fontWeight: 800, color: INK }}>
        [배포]를 누르면 주소가 바로 나오지 않고 이 화면이 먼저 나옵니다
      </div>
      <div
        style={{
          marginTop: 6,
          maxWidth: 520,
          border: `1px solid ${LINE}`,
          borderRadius: 8,
          background: "#fff",
          fontFamily: "Roboto, Arial, sans-serif",
          color: "#3C4043",
          fontSize: 12.5,
          padding: "12px 14px",
        }}
      >
        <div style={{ fontSize: 15, color: "#202124" }}>새 배포</div>
        <div style={{ marginTop: 8 }}>웹 앱에서 내 데이터에 대한 액세스 권한 부여를 나에게 요청합니다.</div>
        <div style={{ marginTop: 10 }}>
          <span
            style={{
              ...ring,
              display: "inline-block",
              background: "#1A73E8",
              color: "#fff",
              fontWeight: 700,
              borderRadius: 4,
              padding: "6px 14px",
            }}
          >
            액세스 승인
          </span>
          <span style={{ color: "#A2452C", fontSize: 11.5, marginLeft: 8 }}>← 누르세요</span>
        </div>
      </div>
      <div style={{ fontSize: 12.5, color: MUTED, lineHeight: 1.6, marginTop: 6 }}>
        <b style={{ color: INK }}>[액세스 승인]</b> → 내 계정 선택 → 위의 ‘Google에서 확인하지 않은 앱’ 화면이{" "}
        <b style={{ color: INK }}>한 번 더</b> 나옵니다. 같은 순서(고급 → 이동 → 모두 선택 → 계속)로 허용하면 주소가
        나옵니다. 설정을 붙여넣을 때 이미 허용했더라도 배포 때 다시 묻는 경우가 있습니다.
      </div>
    </div>
  );
}
// 'Google에서 확인하지 않은 앱' 화면을 실제 모습대로 그린 그림. 캡처 대신 그려서, 개발자 이메일 자리에 누구 주소가 보이는지 설명할 수 있습니다.
function UnverifiedAppGuide() {
  const screen = {
    border: `1px solid ${LINE}`,
    borderRadius: 8,
    background: "#fff",
    padding: "14px 16px 12px",
    color: "#3C4043",
    fontFamily: "Roboto, Arial, sans-serif",
  };
  const ring = {
    display: "inline-block",
    outline: "3px solid #D93025",
    outlineOffset: 3,
    borderRadius: 4,
  };
  const badge = {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    width: 20,
    height: 20,
    borderRadius: 999,
    background: "#D93025",
    color: "#fff",
    fontSize: 12,
    fontWeight: 800,
    flexShrink: 0,
  };
  const label = {
    display: "flex",
    alignItems: "center",
    gap: 6,
    fontSize: 12.5,
    fontWeight: 800,
    color: INK,
    margin: "0 0 6px",
  };
  return (
    <div style={{ marginTop: 8, display: "grid", gap: 10 }}>
      <div>
        <div style={label}>
          <span style={badge}>1</span>
          <span>
            이 화면이 나오면 왼쪽 아래 작은 글씨 <span style={{ color: "#D93025" }}>[고급]</span>(영어: Advanced)을
            누르세요. 파란 버튼은 누르지 마세요.
          </span>
        </div>
        <div style={{ ...screen, maxWidth: 520 }}>
          <svg width="34" height="30" viewBox="0 0 34 30" aria-hidden="true">
            <path d="M17 1 L33 29 L1 29 Z" fill="#DB4437" />
            <rect x="15.5" y="10" width="3" height="10" fill="#fff" />
            <rect x="15.5" y="22.5" width="3" height="3" fill="#fff" />
          </svg>
          <div style={{ fontSize: 19, color: "#202124", margin: "8px 0 6px" }}>Google에서 확인하지 않은 앱</div>
          <div style={{ fontSize: 12.5, lineHeight: 1.6 }}>
            앱에서 Google 계정의 민감한 정보에 대한 액세스를 요청합니다. 개발자(
            <span style={{ color: "#1A73E8", textDecoration: "underline" }}>선생님 이메일</span>)의 앱이 Google에서
            인증을 받기 전에는 앱을 사용하지 마세요.
          </div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 18 }}>
            <span style={ring}>
              <span style={{ fontSize: 12.5, textDecoration: "underline", color: "#3C4043", padding: "0 2px" }}>
                고급
              </span>
            </span>
            <span
              style={{
                background: "#4285F4",
                color: "#fff",
                fontSize: 12,
                fontWeight: 700,
                borderRadius: 4,
                padding: "7px 10px",
                opacity: 0.55,
              }}
            >
              안전한 환경으로 돌아가기
            </span>
          </div>
        </div>
      </div>
      <div>
        <div style={label}>
          <span style={badge}>2</span>
          <span>
            아래에 글이 더 나오면 맨 아래{" "}
            <span style={{ color: "#D93025" }}>[(프로젝트 이름)(으)로 이동(안전하지 않음)]</span>(영어: Go to …
            (unsafe))을 누르세요.
          </span>
        </div>
        <div style={{ ...screen, maxWidth: 520, fontSize: 12.5, lineHeight: 1.6 }}>
          <div style={{ marginBottom: 10 }}>
            계속하려면 위험성을 이해하고 개발자(선생님 이메일)를 신뢰하는 경우에만 계속하세요.
          </div>
          <span style={ring}>
            <span style={{ color: "#3C4043", textDecoration: "underline", padding: "0 2px" }}>
              학생 기초조사(으)로 이동(안전하지 않음)
            </span>
          </span>
        </div>
      </div>
      <div>
        <div style={label}>
          <span style={badge}>3</span>
          <span>
            권한 목록이 나오면 맨 위 <span style={{ color: "#D93025" }}>[모두 선택]</span>(Select all)에 꼭 체크하고
            파란 <span style={{ color: "#D93025" }}>[계속]</span>(Continue)을 누르세요. 하나라도 빠지면 설정 창이나 학생
            화면이 열리지 않습니다.
          </span>
        </div>
        <div style={{ ...screen, maxWidth: 520, fontSize: 12.5, lineHeight: 1.6 }}>
          <div style={{ marginBottom: 8 }}>학생 기초조사에서 액세스할 수 있는 항목을 선택하세요.</div>
          <span style={ring}>
            <span style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "0 4px", fontWeight: 700 }}>
              <span
                style={{
                  width: 14,
                  height: 14,
                  borderRadius: 3,
                  background: "#1A73E8",
                  color: "#fff",
                  fontSize: 11,
                  lineHeight: "14px",
                  textAlign: "center",
                }}
              >
                ✓
              </span>
              모두 선택
            </span>
          </span>
          <div style={{ color: MUTED, margin: "8px 0 0 24px" }}>
            ☑ 이 애플리케이션이 설치된 스프레드시트 보기 및 관리
            <br />☑ Google 애플리케이션 내에서 메시지와 대화상자 표시
            <br />☑ 이메일 주소 보기 …
          </div>
          <div style={{ textAlign: "right", marginTop: 10 }}>
            <span style={ring}>
              <span
                style={{
                  background: "#1A73E8",
                  color: "#fff",
                  fontSize: 12,
                  fontWeight: 700,
                  borderRadius: 4,
                  padding: "5px 12px",
                  display: "inline-block",
                }}
              >
                계속
              </span>
            </span>
          </div>
        </div>
      </div>
      <div style={{ fontSize: 12.3, color: MUTED, lineHeight: 1.6 }}>
        괄호 안의 ‘개발자’ 자리에는 <b style={{ color: INK }}>선생님 본인의 이메일</b>이 보입니다. 사본을 만든 순간 이
        코드는 선생님 것이 되기 때문입니다. ‘확인하지 않은 앱’은 구글 심사를 받지 않았다는 뜻일 뿐이고, 이 코드는{" "}
        <b style={{ color: INK }}>이 시트 하나에만</b> 접근하며 다른 파일·메일·드라이브는 볼 수 없습니다. (이메일 주소
        권한은 설정 창을 시트 주인만 열 수 있는지 확인하는 데만 씁니다.) 권한 허용은 처음 한 번만 하면 됩니다.
      </div>
    </div>
  );
}
function SurveyHelpItem({ title, children }) {
  return (
    <details style={{ borderTop: `1px solid ${LINE}`, padding: "8px 0" }}>
      <summary style={{ cursor: "pointer", fontSize: 12.8, fontWeight: 700, color: INK }}>{title}</summary>
      <div style={{ fontSize: 12.5, color: MUTED, lineHeight: 1.65, marginTop: 6 }}>{children}</div>
    </details>
  );
}
function SurveyInstallStep({ payload, problems, config, setConfig }) {
  const [copyState, setCopyState] = useState(null); // {what: 'code'|'config', ok}
  const [showText, setShowText] = useState(null); // 'code' | 'config' | null
  const blocked = problems.errors.length > 0;
  const copyUrl = templateCopyUrl(SURVEY_TEMPLATE_COPY_URL);
  const method = copyUrl ? "template" : "code";
  const checks = config.installChecks || {};
  const steps = SURVEY_INSTALL_STEPS[method];
  const doneCount = steps.filter((id) => checks[id]).length;
  const code = useMemo(() => (blocked ? "" : buildAppsScriptCode(payload)), [payload, blocked]);
  const configCode = useMemo(() => (blocked ? "" : buildSurveyConfigCode(payload)), [payload, blocked]);
  const stale = !!config.copiedFp && config.copiedFp !== payload.fp;
  const toggle = (id) =>
    setConfig((c) => ({
      ...c,
      installChecks: {
        ...(c.installChecks || {}),
        [id]: !(c.installChecks || {})[id],
      },
    }));
  const markCopied = () =>
    setConfig((c) =>
      c.copiedFp === payload.fp
        ? c
        : {
            ...c,
            copiedFp: payload.fp,
          },
    );
  const doCopy = async (what) => {
    const ok = await copyTextToClipboard(what === "code" ? code : configCode);
    setCopyState({ what, ok });
    if (ok) markCopied();
    else setShowText(what);
  };
  const doSave = () => {
    downloadBlob(
      new Blob([code], {
        type: "text/plain;charset=utf-8",
      }),
      `기초조사_설치코드_${payload.fp}.txt`,
    );
    markCopied();
  };
  const li = { marginBottom: 3 };
  const copyNote = (what) =>
    copyState &&
    copyState.what === what && (
      <span style={{ fontSize: 12.5, color: copyState.ok ? OK : WARN, fontWeight: copyState.ok ? 700 : 400 }}>
        {copyState.ok
          ? `✓ 복사했습니다 (설정 버전 ${payload.fp})`
          : "자동 복사가 막혀 있습니다. 아래 칸을 클릭해 Ctrl+A, Ctrl+C로 복사하세요."}
      </span>
    );
  const copyConfigButton = (
    <button
      disabled={blocked}
      onClick={() => doCopy("config")}
      style={buttonStyle(blocked ? "disabled" : "primary", { padding: "7px 13px" })}
    >
      📋 설정 코드 복사
    </button>
  );
  const textBox = (what) =>
    showText === what &&
    !blocked && (
      <textarea
        readOnly={true}
        value={what === "code" ? code : configCode}
        onFocus={(e) => e.target.select()}
        style={{
          ...inputStyle,
          width: "100%",
          height: 120,
          marginTop: 8,
          fontFamily: "Menlo, Consolas, monospace",
          fontSize: 11,
          whiteSpace: "pre",
        }}
      />
    );
  const deployText = (
    <ol style={{ margin: "4px 0 0", paddingLeft: 18 }}>
      <li style={li}>
        오른쪽 위 파란 <b>[배포] → [새 배포]</b>
      </li>
      <li style={li}>
        ‘유형 선택’ 옆 톱니바퀴 ⚙ → <b>웹 앱</b>
      </li>
      <li style={li}>
        ‘다음 사용자 인증 정보로 실행’은 <b>나</b>, ‘액세스 권한이 있는 사용자’는 <b>모든 사용자</b> → <b>[배포]</b>
      </li>
      <li style={li}>
        <b>[액세스 승인]</b> → 내 계정 선택 → ‘확인하지 않은 앱’ 화면에서 <b>고급 → 이동 → 모두 선택 → 계속</b> (배포
        때도 한 번 더 묻습니다)
      </li>
    </ol>
  );
  // 권한 허용 중 가장 많이 막히는 화면이라 접지 않고 늘 보여줍니다
  const authNote = (
    <div
      style={{ marginTop: 8, border: `1px solid ${WARN}`, background: WARN_BG, borderRadius: 8, padding: "10px 12px" }}
    >
      <div style={{ fontWeight: 800, color: WARN, fontSize: 13 }}>
        ⚠ ‘Google에서 확인하지 않은 앱’ 화면이 나와도 멈추지 마세요
      </div>
      <div style={{ fontSize: 12.3, color: WARN, marginTop: 2 }}>
        설정을 처음 저장할 때와, 배포 창에서 <b>[액세스 승인]</b>을 누른 뒤에 나옵니다. 영어로 나오면 ‘Google hasn’t
        verified this app’입니다.
      </div>
      <UnverifiedAppGuide />
    </div>
  );
  // 템플릿 링크를 못 쓸 때의 예비 방법 (새 시트에 설치 코드를 직접 붙여넣기)
  const codeSteps = (
    <div>
      <InstallCheckItem no={1} done={checks.c1} onToggle={() => toggle("c1")} title="새 구글 시트 만들기">
        조사에 쓸 구글 계정으로 로그인한 상태에서 아래 버튼을 누르면 빈 시트가 바로 열립니다. 왼쪽 위 ‘제목 없는
        스프레드시트’를 눌러 이름을 붙이세요. (예: 2025 입학생 기초조사 응답)
        <div style={{ marginTop: 6 }}>
          <a
            href="https://sheets.new"
            target="_blank"
            rel="noopener noreferrer"
            style={{ ...buttonStyle("ghost", { padding: "7px 13px" }), textDecoration: "none" }}
          >
            ➕ 새 구글 시트 열기 ↗
          </a>
        </div>
      </InstallCheckItem>
      <InstallCheckItem no={2} done={checks.c2} onToggle={() => toggle("c2")} title="Apps Script 열기">
        시트 위쪽 메뉴에서 <b>[확장 프로그램 → Apps Script]</b>를 누릅니다. 새 탭에 코드 편집기가 열립니다.
      </InstallCheckItem>
      <InstallCheckItem no={3} done={checks.c3} onToggle={() => toggle("c3")} title="설치 코드 붙여넣고 저장">
        편집기에 있는 글(function myFunction…)을 <b>모두 지우고</b>, <b>[설치 코드 복사]</b>를 누른 뒤 붙여넣고{" "}
        <b>저장(Ctrl+S)</b>합니다. 지금 조사 설정이 코드에 함께 들어 있습니다.
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center", marginTop: 6 }}>
          <button
            disabled={blocked}
            onClick={() => doCopy("code")}
            style={buttonStyle(blocked ? "disabled" : "primary", { padding: "7px 13px" })}
          >
            📋 설치 코드 복사
          </button>
          <button
            disabled={blocked}
            onClick={doSave}
            style={buttonStyle(blocked ? "disabled" : "ghost", { padding: "7px 13px" })}
          >
            <Download size={13} /> 코드를 파일로 저장(.txt)
          </button>
          {!blocked && (
            <button
              onClick={() => setShowText((v) => (v === "code" ? null : "code"))}
              style={{ ...buttonStyle("ghost"), border: "none", color: MUTED }}
            >
              {showText === "code" ? "코드 숨기기" : "코드 직접 보기"}
            </button>
          )}
          {copyNote("code")}
        </div>
        {textBox("code")}
      </InstallCheckItem>
      <InstallCheckItem no={4} done={checks.c4} onToggle={() => toggle("c4")} title="웹 앱으로 배포하기">
        {deployText}
        <DeployDialogGuide />
        <DeployAuthPicture />
        처음 한 번은 <b>[액세스 승인]</b> → 내 계정 선택 → ‘확인하지 않은 앱’ 화면에서 <b>[고급]</b> →{" "}
        <b>(프로젝트 이름)(으)로 이동</b> → <b>[모두 선택]</b> 체크 → <b>[계속]</b> 순서로 이어집니다. (아래 그림)
        {authNote}
      </InstallCheckItem>
      <InstallCheckItem no={5} done={checks.c5} onToggle={() => toggle("c5")} title="웹 앱 주소 복사하기">
        배포가 끝나면 ‘배포가 업데이트되었습니다’ 화면이 나옵니다. <b>아래쪽 ‘웹 앱 → URL’ 밑의 [복사]</b>를 눌러 (위쪽
        ‘배포 ID’ 옆 [복사]가 아닙니다) 아래 5번 칸에 붙여넣습니다.
        <DeployDonePicture />
      </InstallCheckItem>
      {!stale && (
        <div style={{ fontSize: 12.3, color: MUTED, lineHeight: 1.6, margin: "4px 0 8px" }}>
          설치한 뒤에 조사 설정이 바뀌면 코드를 다시 붙여넣지 말고, <b>[설정 코드 복사]</b> → 시트 메뉴{" "}
          <b>[📋 기초조사 → 설정 붙여넣기]</b>만 하면 됩니다. 다시 배포할 필요가 없습니다.
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 6 }}>
            {copyConfigButton}
            {copyNote("config")}
          </div>
          {textBox("config")}
        </div>
      )}
    </div>
  );
  return (
    <SurveyStep
      no={4}
      title="구글 시트에 설치하기 (처음 한 번)"
      desc="학생이 낸 내용은 선생님의 구글 계정에 있는 시트 하나에만 저장됩니다. 이 프로그램이나 다른 곳으로는 보내지지 않습니다."
    >
      {problems.errors.length > 0 && (
        <div
          style={{
            background: WARN_BG,
            color: WARN,
            borderRadius: 8,
            padding: "9px 12px",
            fontSize: 12.8,
            lineHeight: 1.6,
            marginBottom: 10,
          }}
        >
          {problems.errors.map((e, i) => (
            <div key={i}>⚠ {e}</div>
          ))}
        </div>
      )}
      {problems.warnings.length > 0 && (
        <div
          style={{
            fontSize: 12.3,
            color: MUTED,
            lineHeight: 1.6,
            marginBottom: 10,
          }}
        >
          {problems.warnings.map((w, i) => (
            <div key={i}>· {w}</div>
          ))}
        </div>
      )}
      {stale && (
        <div
          style={{
            background: WARN_BG,
            color: WARN,
            borderRadius: 8,
            padding: "10px 12px",
            fontSize: 12.8,
            lineHeight: 1.65,
            marginBottom: 12,
          }}
        >
          <div style={{ fontWeight: 800, marginBottom: 4 }}>
            ⚠ 시트에 넘긴 뒤 조사 설정(학기·반별 인원·안내·편제표)이 바뀌었습니다.
          </div>
          이미 설치했다면 <b>[설정 코드 복사]</b>를 누르고, 구글 시트 위 메뉴 <b>[📋 기초조사 → 설정 붙여넣기]</b>에
          붙여넣어 저장하세요. <b>다시 배포할 필요가 없고</b> 학생 주소(QR 코드)도 그대로입니다.
          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 8 }}>
            {copyConfigButton}
            {copyNote("config")}
          </div>
          <div style={{ fontSize: 12, marginTop: 8, color: MUTED }}>
            시트에 📋 기초조사 메뉴가 없다면 이 프로그램의 예전 버전으로 설치한 시트입니다. 이번 한 번만 Apps Script에
            새 설치 코드를 붙여넣고 저장한 뒤 <b>배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포</b>를 누르세요.
            다음부터는 메뉴로 바꿀 수 있습니다.
          </div>
        </div>
      )}
      {copyUrl ? (
        <div>
          <div style={{ fontSize: 12.3, color: MUTED, marginBottom: 8 }}>
            세 단계입니다. 하나씩 끝낼 때마다 왼쪽 네모에 체크하세요. ({doneCount}/{steps.length} 완료)
          </div>
          <InstallCheckItem no={1} done={checks.t1} onToggle={() => toggle("t1")} title="템플릿 시트 사본 만들기">
            조사에 쓸 구글 계정으로 로그인한 상태에서 아래 버튼을 누르고 <b>[사본 만들기]</b>를 누릅니다. 만들어진 시트
            이름을 알아보기 쉽게 바꿔 두세요. (예: 2025 입학생 기초조사 응답)
            <div style={{ marginTop: 6 }}>
              <a
                href={copyUrl}
                target="_blank"
                rel="noopener noreferrer"
                style={{ ...buttonStyle("primary", { padding: "7px 13px" }), textDecoration: "none" }}
              >
                📄 템플릿 사본 만들기 ↗
              </a>
            </div>
          </InstallCheckItem>
          <InstallCheckItem
            no={2}
            done={checks.t2}
            onToggle={() => toggle("t2")}
            title="시트 안에서 설정 넣고 배포하기"
          >
            사본 시트가 열리면 첫 화면(시작하기)에 같은 안내가 적혀 있습니다.
            <ol style={{ margin: "4px 0 0", paddingLeft: 18 }}>
              <li style={li}>
                여기서 <b>[설정 코드 복사]</b>를 누릅니다.
              </li>
              <li style={li}>
                사본 시트 위 메뉴 <b>[📋 기초조사 → 설정 붙여넣기 / 바꾸기]</b>를 열어 Ctrl+V로 붙여넣고 <b>[저장]</b>
                합니다. 메뉴는 시트가 열리고 몇 초 뒤에 생깁니다. 안 보이면 새로고침하세요.
              </li>
              <li style={li}>
                저장한 창에 <b>배포 안내와 [Apps Script 편집기 열기] 버튼</b>이 이어서 나옵니다. 버튼으로 편집기를 열고
                안내대로 <b>[배포] → [새 배포] → ⚙ 웹 앱 → [배포]</b>를 누릅니다. (아래 그림과 같은 안내가 시트 창에도
                나옵니다)
              </li>
            </ol>
            <SheetMenuPicture />
            처음 한 번은 ‘승인 필요’ 창이 뜹니다 → <b>[계속]</b> → 내 계정 선택 → ‘확인하지 않은 앱’ 화면에서{" "}
            <b>[고급]</b> → <b>(프로젝트 이름)(으)로 이동</b> → <b>[모두 선택]</b> 체크 → <b>[계속]</b>.
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", marginTop: 6 }}>
              {copyConfigButton}
              {copyNote("config")}
            </div>
            {textBox("config")}
            {authNote}
            <div style={{ marginTop: 10, fontWeight: 800, fontSize: 13 }}>배포 창은 이렇게 생겼습니다</div>
            <div style={{ fontSize: 12.5, color: MUTED, marginTop: 2 }}>
              시트 창에는 순서만 짧게 적혀 있습니다. 화면이 낯설면 아래 그림과 맞춰 보세요.
            </div>
            {deployText}
            <DeployDialogGuide />
            <DeployAuthPicture />
          </InstallCheckItem>
          <InstallCheckItem
            no={3}
            done={checks.t3}
            onToggle={() => toggle("t3")}
            title="웹 앱 주소를 아래 5번 칸에 붙여넣기"
          >
            배포가 끝나면 ‘배포가 업데이트되었습니다’ 화면이 나옵니다. <b>아래쪽 ‘웹 앱 → URL’ 밑의 [복사]</b>를 눌러
            (위쪽 ‘배포 ID’ 옆 [복사]가 아닙니다) 아래 5번 칸에 붙여넣습니다. 이 주소는 나중에 설정을 바꿔도 그대로이고,
            다시 배포할 필요가 없습니다.
            <DeployDonePicture />
          </InstallCheckItem>
          <details style={{ marginTop: 4 }}>
            <summary style={{ cursor: "pointer", color: MUTED, fontSize: 12.3 }}>
              다른 방법: 템플릿 링크가 열리지 않을 때 새 시트에 코드 붙여넣기
            </summary>
            <div style={{ marginTop: 8 }}>{codeSteps}</div>
          </details>
        </div>
      ) : (
        <div>
          <div style={{ fontSize: 12.3, color: MUTED, marginBottom: 8 }}>
            하나씩 끝낼 때마다 왼쪽 네모에 체크하세요. ({doneCount}/{steps.length} 완료)
          </div>
          {codeSteps}
        </div>
      )}
      <div style={{ marginTop: 6 }}>
        <SurveyHelpItem title="학교(교육청) 구글 계정이라 ‘모든 사용자’가 목록에 없어요">
          학교 구글 워크스페이스 계정은 관리자 설정 때문에 ‘모든 사용자’를 고를 수 없을 수 있습니다. 이때는{" "}
          <b>‘(학교 도메인) 내 모든 사용자’</b>로 배포하고, 학생들이 <b>학교 구글 계정으로 로그인한 휴대폰</b>에서 QR
          코드를 열게 안내하세요. 학생이 개인 계정만 쓴다면 개인 구글 계정으로 시트를 만들어 설치하는 편이 쉽습니다.
        </SurveyHelpItem>
        <SurveyHelpItem title="학생이 QR 코드로 열었더니 구글 로그인 화면이 나와요">
          배포할 때 ‘액세스 권한이 있는 사용자’를 <b>모든 사용자</b>로 하지 않은 경우입니다. Apps Script 편집기에서{" "}
          <b>배포 → 배포 관리 → 연필(수정)</b>을 눌러 액세스를 ‘모든 사용자’로 바꾸고 <b>[배포]</b>하세요. 주소는
          그대로입니다.
        </SurveyHelpItem>
        <SurveyHelpItem title="학생 화면에 ‘아직 조사 준비 중이에요’가 나와요">
          시트에 조사 설정이 아직 없습니다. <b>[설정 코드 복사]</b>를 누르고 시트 메뉴{" "}
          <b>[📋 기초조사 → 설정 붙여넣기]</b>에 붙여넣어 저장하세요. 저장하면 바로 학생 화면이 열립니다.
        </SurveyHelpItem>
        <SurveyHelpItem title="시트 위에 📋 기초조사 메뉴가 안 보여요">
          시트를 연 뒤 몇 초 기다리거나 새로고침(F5)하세요. 코드를 방금 붙여넣었다면 시트 탭을 새로고침해야 메뉴가
          생깁니다. 그래도 없으면 Apps Script 편집기에서 코드가 저장되었는지(파일 이름 옆에 저장 안 됨 표시가 없는지)
          확인하세요.
        </SurveyHelpItem>
      </div>
    </SurveyStep>
  );
}
function SurveyQrStep({ config, setConfig, payload, onPrintQr }) {
  const url = checkWebAppUrl(config.webAppUrl);
  const firstClass = payload.classes.length ? payload.classes[0].c : 1;
  return (
    <SurveyStep
      no={5}
      title="QR 코드 나눠주기"
      desc="4번의 웹 앱 URL을 붙여넣으면 반마다 QR 코드 안내문을 인쇄할 수 있습니다. 담임 선생님이 교실에 붙이거나 나눠주세요."
    >
      <input
        value={config.webAppUrl}
        onChange={(e) =>
          setConfig((c) => ({
            ...c,
            webAppUrl: e.target.value,
          }))
        }
        placeholder="https://script.google.com/macros/s/…/exec"
        style={{
          ...inputStyle,
          width: "100%",
          fontSize: 13,
          padding: "9px 11px",
          borderColor: url.ok || url.empty ? LINE : WARN,
        }}
      />
      {!url.ok && !url.empty && (
        <div
          style={{
            fontSize: 12.5,
            color: WARN,
            marginTop: 6,
            lineHeight: 1.55,
          }}
        >
          {url.message}
        </div>
      )}
      {url.ok && (
        <div
          style={{
            display: "flex",
            gap: 16,
            alignItems: "center",
            flexWrap: "wrap",
            marginTop: 12,
          }}
        >
          <div
            style={{
              border: `1px solid ${LINE}`,
              borderRadius: 8,
              padding: 6,
              background: "#fff",
            }}
          >
            <QrSvg text={surveyClassUrl(url.url, firstClass)} size={132} />
          </div>
          <div
            style={{
              flex: 1,
              minWidth: 260,
              fontSize: 12.8,
              color: MUTED,
              lineHeight: 1.65,
            }}
          >
            <div
              style={{
                color: OK,
                fontWeight: 700,
                marginBottom: 4,
              }}
            >
              {"\u2713 \uC62C\uBC14\uB978 \uC6F9 \uC571 \uC8FC\uC18C\uC785\uB2C8\uB2E4. (\uC67C\uCABD\uC740 "}
              {firstClass}반용 QR 코드)
            </div>
            {"\uBC18\uB9C8\uB2E4 \uC8FC\uC18C \uB05D\uC5D0 "}
            <b
              style={{
                color: INK,
              }}
            >
              ?ban=반
            </b>
            이 붙어서, QR 코드로 들어온 학생은 자기 반이 미리 골라져 있습니다. 인쇄하기 전에 휴대폰으로 한 번 찍어 학생
            화면이 열리는지 확인하세요.
            <div
              style={{
                display: "flex",
                gap: 8,
                flexWrap: "wrap",
                marginTop: 8,
              }}
            >
              <button
                disabled={payload.classes.length === 0}
                onClick={() => onPrintQr(url.url)}
                style={buttonStyle(payload.classes.length ? "primary" : "disabled", {
                  padding: "8px 14px",
                })}
              >
                🖨 반별 QR 안내문 인쇄 ({payload.classes.length}장)
              </button>
              <a
                href={surveyClassUrl(url.url, firstClass)}
                target="_blank"
                rel="noopener noreferrer"
                style={{
                  ...buttonStyle("ghost", {
                    padding: "8px 14px",
                  }),
                  textDecoration: "none",
                }}
              >
                주소 열어 확인
              </a>
            </div>
          </div>
        </div>
      )}
    </SurveyStep>
  );
}
function SurveyFileRow({ f, payload, onRemove, onToggleDisabled }) {
  const staleMeta = f.status === "ok" && f.meta && f.meta.fp && f.meta.fp !== payload.fp;
  return (
    <div
      style={{
        background: PAPER,
        borderRadius: 8,
        padding: "8px 10px",
        fontSize: 12.5,
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          gap: 8,
        }}
      >
        {f.status === "ok" ? (
          <CheckCircle2
            size={14}
            color={OK}
            style={{
              flexShrink: 0,
            }}
          />
        ) : (
          <AlertCircle
            size={14}
            color={WARN}
            style={{
              flexShrink: 0,
            }}
          />
        )}
        <span
          style={{
            flex: 1,
            minWidth: 0,
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
          }}
        >
          {f.name}
        </span>
        {f.status === "ok" && (
          <span
            style={{
              color: MUTED,
              flexShrink: 0,
            }}
          >
            {"\u00B7 \uC81C\uCD9C "}
            {f.submissions.length}건
          </span>
        )}
        {f.status === "ok" && onToggleDisabled && (
          <label
            title="끄면 이 파일의 응답은 점검에 쓰지 않습니다"
            style={{
              display: "flex",
              alignItems: "center",
              gap: 3,
              fontSize: 11.5,
              color: f.disabled ? WARN : MUTED,
              flexShrink: 0,
              cursor: "pointer",
            }}
          >
            <input type="checkbox" checked={!f.disabled} onChange={() => onToggleDisabled(f.id)} />
            계산에 사용
          </label>
        )}
        <button
          onClick={() => onRemove(f.id)}
          style={{
            border: "none",
            background: "none",
            cursor: "pointer",
            color: MUTED,
            flexShrink: 0,
          }}
        >
          <X size={14} />
        </button>
      </div>
      {f.status === "error" && (
        <div
          style={{
            color: WARN,
            marginTop: 4,
            lineHeight: 1.5,
          }}
        >
          {f.error}
        </div>
      )}
      {f.status === "ok" && f.disabled && (
        <div
          style={{
            color: WARN,
            marginTop: 4,
            lineHeight: 1.5,
          }}
        >
          ‘계산에 사용’이 꺼져 있어 이 응답 파일은 점검과 제출 현황에 쓰지 않습니다.
        </div>
      )}
      {f.status === "ok" && f.meta && (
        <div
          style={{
            color: MUTED,
            marginTop: 3,
          }}
        >
          {"\uC870\uC0AC: "}
          {f.meta.title || "(제목 없음)"}
          {" \u00B7 \uCF54\uB4DC \uBC84\uC804 "}
          {f.meta.fp}
        </div>
      )}
      {f.status === "ok" && !f.meta && (
        <div
          style={{
            color: MUTED,
            marginTop: 3,
          }}
        >
          설정 시트의 조사 정보가 없는 파일입니다. (선택과목 시트만으로 읽었습니다)
        </div>
      )}
      {staleMeta && (
        <div
          style={{
            color: WARN,
            marginTop: 4,
            lineHeight: 1.5,
          }}
        >
          ⚠ 이 응답은 지금 설정과 다른 설정 버전으로 받았습니다. 조사 도중 편제표나 조사 범위를 바꿨다면 과목이 맞는지 ③
          점검 화면에서 확인하세요.
        </div>
      )}
      {f.status === "ok" && f.skipped > 0 && (
        <div
          style={{
            color: WARN,
            marginTop: 4,
          }}
        >
          {"\u26A0 \uC77D\uC9C0 \uBABB\uD55C \uC904 "}
          {f.skipped}개 (반·번호·이름·학기·과목 중 빈칸이 있는 줄)
        </div>
      )}
    </div>
  );
}
function SurveyResponsesStep({
  surveyFiles,
  onFiles,
  onFetchLive,
  surveyFetch,
  onRemove,
  data,
  payload,
  config,
  choice,
  setChoice,
  excluded,
  toggleExcluded,
  onGoTransfer,
  onToggleDisabled,
}) {
  const total = config.classes.reduce(
    (a, k) => a + (k.size || 0) - (k.skip || []).filter((n) => n <= (k.size || 0)).length,
    0,
  );
  const okFiles = surveyFiles.filter((f) => f.status === "ok" && !f.disabled);
  const issueCount = data.conflicts.length + data.sameName.length + data.outside.length;
  const liveUrl = checkWebAppUrl(config.webAppUrl);
  const liveFile = surveyFiles.find((f) => f.id === "sf-live");
  const box = {
    background: "#fff",
    border: `1px solid ${LINE}`,
    borderRadius: 10,
    padding: 14,
    marginTop: 12,
  };
  const subLine = (s) =>
    `${s.name} · ${s.at || "시각 미상"} 제출 · 과목 ${s.subjects.length}개${s.away.length ? ` · 다른 학교 ${s.away.map(formatSemester).join(", ")}` : ""}`;
  return (
    <SurveyStep
      no={6}
      title="응답 가져오기"
      desc="학생들이 낸 응답을 구글 시트에서 바로 가져옵니다. 조사 중간에 여러 번 눌러도 됩니다 (같은 제출은 겹쳐 세지 않습니다)."
    >
      <div
        style={{
          border: `1px solid ${liveUrl.ok ? ACCENT : LINE}`,
          background: liveUrl.ok ? ACCENT_BG : "#fff",
          borderRadius: 10,
          padding: "12px 14px",
        }}
      >
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <button
            disabled={!liveUrl.ok || (surveyFetch && surveyFetch.busy)}
            onClick={onFetchLive}
            style={buttonStyle(liveUrl.ok && !(surveyFetch && surveyFetch.busy) ? "primary" : "disabled", {
              padding: "8px 14px",
            })}
          >
            {surveyFetch && surveyFetch.busy ? "가져오는 중…" : "🔄 구글 시트에서 지금 가져오기"}
          </button>
          {liveUrl.ok && (
            <a
              href={`${liveUrl.url}?export=${encodeURIComponent(config.exportKey || "")}`}
              target="_blank"
              rel="noopener noreferrer"
              style={{ ...buttonStyle("ghost", { padding: "8px 12px" }), textDecoration: "none" }}
              title="시트가 돌려주는 내용을 새 탭에서 봅니다 (선생님 확인용 · 이 주소는 학생에게 주지 마세요)"
            >
              주소 열어 확인
            </a>
          )}
          <span style={{ fontSize: 12.5, color: MUTED, lineHeight: 1.55 }}>
            {liveUrl.ok
              ? liveFile
                ? `마지막으로 가져온 때: ${liveFile.name.replace(/^구글 시트에서 가져옴 \(|\)$/g, "")} · 제출 ${liveFile.submissions.length}건`
                : "5번 칸의 웹 앱 주소로 시트에 저장된 응답을 받아옵니다. 시트를 열거나 내려받을 필요가 없습니다."
              : "먼저 5번 칸에 웹 앱 주소를 넣어 주세요."}
          </span>
        </div>
        {surveyFetch && surveyFetch.error && (
          <div style={{ fontSize: 12.5, color: WARN, marginTop: 8, lineHeight: 1.55 }}>⚠ {surveyFetch.error}</div>
        )}
      </div>
      <details style={{ marginTop: 10 }}>
        <summary style={{ cursor: "pointer", fontSize: 12.5, color: MUTED }}>
          다른 방법: 시트를 엑셀로 내려받아 올리기 (가져오기가 안 될 때)
        </summary>
        <div style={{ fontSize: 12.5, color: MUTED, lineHeight: 1.6, margin: "6px 0 8px" }}>
          조사용 구글 시트에서 <b style={{ color: INK }}>파일 → 다운로드 → Microsoft Excel(.xlsx)</b>로 내려받아 여기에
          올립니다. 여러 번 올려도 같은 제출은 겹쳐 세지 않습니다.
        </div>
        <RegFileDropzone onFiles={onFiles} label="내려받은 응답 파일(.xlsx)을 끌어다 놓거나" />
      </details>
      {surveyFiles.length > 0 && (
        <div
          style={{
            marginTop: 10,
            display: "flex",
            flexDirection: "column",
            gap: 6,
          }}
        >
          {surveyFiles.map((f) => (
            <SurveyFileRow key={f.id} f={f} payload={payload} onRemove={onRemove} onToggleDisabled={onToggleDisabled} />
          ))}
        </div>
      )}
      {okFiles.length > 0 && (
        <>
          <div
            style={{
              display: "flex",
              gap: 10,
              flexWrap: "wrap",
              marginTop: 14,
            }}
          >
            <StatCard label="제출한 학생" value={data.records.length} tone={ACCENT} />
            <StatCard label="명단 인원" value={total} />
            <StatCard label="아직 안 낸 학생" value={data.missingTotal} tone={data.missingTotal > 0 ? WARN : OK} />
            <StatCard label="확인할 것" value={issueCount} tone={issueCount > 0 ? WARN : OK} />
          </div>
          {data.missingByClass.length > 0 && (
            <div
              style={{
                ...box,
                padding: 0,
                overflow: "hidden",
              }}
            >
              <div
                style={{
                  padding: "12px 14px 6px",
                  fontSize: 13.5,
                  fontWeight: 800,
                }}
              >
                반별 제출 현황
              </div>
              <div
                style={{
                  overflowX: "auto",
                }}
              >
                <table
                  style={{
                    borderCollapse: "collapse",
                    width: "100%",
                    fontSize: 12.5,
                  }}
                >
                  <thead>
                    <tr
                      style={{
                        background: PAPER,
                      }}
                    >
                      <Th>반</Th>
                      <Th>인원</Th>
                      <Th>제출</Th>
                      <Th>아직 안 낸 번호</Th>
                    </tr>
                  </thead>
                  <tbody>
                    {data.missingByClass.map((k) => (
                      <tr
                        key={k.c}
                        style={{
                          borderTop: `1px solid ${LINE}`,
                        }}
                      >
                        <Td
                          style={{
                            fontWeight: 700,
                          }}
                        >
                          {payload.grade}
                          {"\uD559\uB144 "}
                          {k.c}반
                        </Td>
                        <Td
                          style={{
                            whiteSpace: "normal",
                          }}
                        >
                          {k.size}명{k.skip && k.skip.length ? ` (결번 ${k.skip.join(", ")}번 제외)` : ""}
                        </Td>
                        <Td
                          style={{
                            color: k.missing.length ? INK : OK,
                          }}
                        >
                          {k.submitted}명
                        </Td>
                        <Td
                          style={{
                            color: k.missing.length ? WARN : OK,
                            whiteSpace: "normal",
                            fontWeight: k.missing.length ? 700 : 400,
                          }}
                        >
                          {k.missing.length ? k.missing.map((n) => `${n}번`).join(", ") : "모두 냈습니다 ✓"}
                        </Td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          {data.conflicts.length > 0 && (
            <div
              style={{
                ...box,
                borderColor: "#E8C9BB",
                background: WARN_BG,
              }}
            >
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                  color: WARN,
                  marginBottom: 4,
                }}
              >
                같은 번호로 다른 이름이 제출됐습니다 ({data.conflicts.length}건)
              </div>
              <div
                style={{
                  fontSize: 12.5,
                  color: WARN,
                  lineHeight: 1.55,
                  marginBottom: 8,
                }}
              >
                {
                  "\uBC88\uD638\uB97C \uC798\uBABB \uB20C\uB800\uAC70\uB098 \uB2E4\uB978 \uD559\uC0DD\uC774 \uB300\uC2E0 \uB0C8\uC744 \uC218 \uC788\uC2B5\uB2C8\uB2E4. \uD559\uC0DD\uC5D0\uAC8C \uD655\uC778\uD55C \uB4A4 "
                }
                <b>맞는 제출</b>을 골라주세요. (처음에는 마지막 제출이 골라져 있습니다. 번호를 잘못 낸 학생은 자기
                번호로 다시 내게 하세요)
              </div>
              {data.conflicts.map((c) => (
                <div
                  key={c.id}
                  style={{
                    background: "#fff",
                    borderRadius: 8,
                    padding: "8px 10px",
                    marginBottom: 6,
                  }}
                >
                  <div
                    style={{
                      fontSize: 12.5,
                      fontWeight: 800,
                      marginBottom: 4,
                    }}
                  >
                    {c.idText}
                  </div>
                  {c.subs.map((s) => (
                    <label
                      key={s.key}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 7,
                        fontSize: 12.5,
                        cursor: "pointer",
                        padding: "2px 0",
                      }}
                    >
                      <input
                        type="radio"
                        name={`svc-${c.id}`}
                        checked={c.pickedKey === s.key}
                        onChange={() => setChoice(c.id, s.key)}
                      />
                      {subLine(s)}
                    </label>
                  ))}
                </div>
              ))}
            </div>
          )}
          {data.sameName.length > 0 && (
            <div style={box}>
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                  marginBottom: 4,
                }}
              >
                같은 이름이 다른 번호로 제출됐습니다 ({data.sameName.length}명)
              </div>
              <div
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  lineHeight: 1.55,
                  marginBottom: 8,
                }}
              >
                {
                  "\uB3D9\uBA85\uC774\uC778\uC774\uBA74 \uADF8\uB300\uB85C \uB450\uC138\uC694 (\uC11C\uB85C \uB2E4\uB978 \uD559\uC0DD\uC73C\uB85C \uC810\uAC80\uB429\uB2C8\uB2E4). \uD55C \uD559\uC0DD\uC774 \uBC88\uD638\uB97C \uC798\uBABB \uB20C\uB7EC \uB450 \uBC88 \uB0B8 \uAC83\uC774\uBA74 "
                }
                <b
                  style={{
                    color: INK,
                  }}
                >
                  잘못된 쪽을 [빼기]
                </b>
                {" \uD558\uC138\uC694."}
              </div>
              {data.sameName.map((g) => (
                <div
                  key={g.name}
                  style={{
                    background: PAPER,
                    borderRadius: 8,
                    padding: "8px 10px",
                    marginBottom: 6,
                  }}
                >
                  {g.records.map((r) => (
                    <div
                      key={r.uid}
                      style={{
                        display: "flex",
                        alignItems: "center",
                        gap: 8,
                        fontSize: 12.5,
                        padding: "2px 0",
                        flexWrap: "wrap",
                      }}
                    >
                      <b>{r.idText}</b>
                      <span
                        style={{
                          color: MUTED,
                        }}
                      >
                        {r.name}
                        {" \u00B7 "}
                        {r.submittedAt}
                        {" \uC81C\uCD9C"}
                      </span>
                      <button
                        onClick={() => toggleExcluded(r.submitKey)}
                        style={buttonStyle("danger", {
                          fontSize: 11.5,
                          padding: "3px 8px",
                        })}
                      >
                        빼기
                      </button>
                    </div>
                  ))}
                </div>
              ))}
            </div>
          )}
          {data.outside.length > 0 && (
            <div style={box}>
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                  marginBottom: 4,
                }}
              >
                명단(반별 인원) 밖의 번호로 제출됐습니다 ({data.outside.length}명)
              </div>
              <div
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  lineHeight: 1.55,
                  marginBottom: 6,
                }}
              >
                2번의 반별 인원이 맞는지, 다른 학년 조사 파일이 섞이지 않았는지 확인하세요. 점검에는 포함됩니다.
              </div>
              {data.outside.map((r) => (
                <div
                  key={r.uid}
                  style={{
                    fontSize: 12.5,
                    padding: "2px 0",
                  }}
                >
                  {"\u00B7 "}
                  {r.idText} {r.name}
                </div>
              ))}
            </div>
          )}
          {data.excludedSubs.length > 0 && (
            <div style={box}>
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                  marginBottom: 6,
                }}
              >
                뺀 제출 ({data.excludedSubs.length}건) — 점검에 쓰지 않습니다
              </div>
              {data.excludedSubs.map((s) => (
                <div
                  key={s.key}
                  style={{
                    display: "flex",
                    alignItems: "center",
                    gap: 8,
                    fontSize: 12.5,
                    padding: "2px 0",
                    flexWrap: "wrap",
                  }}
                >
                  <span>
                    {surveyIdText(s)}
                    {" \u00B7 "}
                    {subLine(s)}
                  </span>
                  <button
                    onClick={() => toggleExcluded(s.key)}
                    style={buttonStyle("ghost", {
                      fontSize: 11.5,
                      padding: "3px 8px",
                    })}
                  >
                    다시 넣기
                  </button>
                </div>
              ))}
            </div>
          )}
          {data.transferMarked.length > 0 && (
            <div style={box}>
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                  marginBottom: 4,
                }}
              >
                ‘다른 학교에 다녔어요’를 표시한 학생 ({data.transferMarked.length}명)
              </div>
              <div
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                  lineHeight: 1.55,
                  marginBottom: 6,
                }}
              >
                {"\uC804\uC785\uC0DD\uC77C \uAC00\uB2A5\uC131\uC774 \uB192\uC2B5\uB2C8\uB2E4. "}
                <b
                  style={{
                    color: INK,
                  }}
                >
                  ⑤ 전입생
                </b>
                {
                  " \uD0ED\uC5D0\uC11C \uBCF8\uAD50 \uCCAB \uD559\uAE30\uC640 \uC804\uC801\uAD50 \uC774\uC218 \uACFC\uBAA9\uC744 \uC785\uB825\uD574\uC57C \uD559\uC810\uC774 \uB9DE\uAC8C \uACC4\uC0B0\uB429\uB2C8\uB2E4."
                }
              </div>
              {data.transferMarked.map((r) => (
                <div
                  key={r.uid}
                  style={{
                    fontSize: 12.5,
                    padding: "2px 0",
                  }}
                >
                  {"\u00B7 "}
                  {r.idText} {r.name}
                  {" \u2014 \uB2E4\uB978 \uD559\uAD50: "}
                  {r.away.map(formatSemester).join(", ")}
                </div>
              ))}
              <button
                onClick={onGoTransfer}
                style={buttonStyle("soft", {
                  marginTop: 8,
                })}
              >
                {"\u2464 \uC804\uC785\uC0DD \uD0ED\uC73C\uB85C \uAC00\uAE30 "}
                <ArrowRight size={13} />
              </button>
            </div>
          )}
          {data.memos.length > 0 && (
            <div style={box}>
              <div
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                  marginBottom: 6,
                }}
              >
                학생이 남긴 말 ({data.memos.length}건)
              </div>
              {data.memos.map((r) => (
                <div
                  key={r.uid}
                  style={{
                    fontSize: 12.5,
                    padding: "3px 0",
                    lineHeight: 1.5,
                  }}
                >
                  <b>
                    {r.idText} {r.name}
                  </b>{" "}
                  <span
                    style={{
                      color: MUTED,
                    }}
                  >
                    {"\u2014 "}
                    {r.memo}
                  </span>
                </div>
              ))}
            </div>
          )}
          {data.resubmitted.length > 0 && (
            <details
              style={{
                ...box,
                fontSize: 12.5,
                color: MUTED,
              }}
            >
              <summary
                style={{
                  cursor: "pointer",
                }}
              >
                {"\uC5EC\uB7EC \uBC88 \uB0B8 \uD559\uC0DD "}
                {data.resubmitted.length}명 — 마지막 제출을 씁니다 (눌러서 보기)
              </summary>
              <div
                style={{
                  marginTop: 6,
                  lineHeight: 1.6,
                }}
              >
                {data.resubmitted.map((x) => `${x.idText} ${x.name}(${x.count}번)`).join(", ")}
              </div>
            </details>
          )}
        </>
      )}
    </SurveyStep>
  );
}
function SurveyPanel(p) {
  return (
    <div>
      <SurveyScopeStep
        config={p.config}
        setConfig={p.setConfig}
        poolDefs={p.poolDefs}
        payload={p.payload}
        admissionYear={p.admissionYear}
      />
      <SurveyRosterStep config={p.config} setConfig={p.setConfig} />
      <SurveyTextStep
        config={p.config}
        setConfig={p.setConfig}
        onPreview={p.onPreview}
        canPreview={
          p.hasCurriculum &&
          !!p.config.grade &&
          !!p.config.term &&
          p.payload.semesters.length > 0 &&
          p.payload.classes.length > 0
        }
      />
      <SurveyInstallStep payload={p.payload} problems={p.problems} config={p.config} setConfig={p.setConfig} />
      <SurveyQrStep config={p.config} setConfig={p.setConfig} payload={p.payload} onPrintQr={p.onPrintQr} />
      <SurveyResponsesStep
        surveyFiles={p.surveyFiles}
        onFiles={p.onSurveyFiles}
        onFetchLive={p.onFetchLive}
        surveyFetch={p.surveyFetch}
        onRemove={p.removeSurveyFile}
        data={p.surveyData}
        payload={p.payload}
        config={p.config}
        choice={p.surveyChoice}
        setChoice={p.setSurveyChoice}
        excluded={p.surveyExcluded}
        toggleExcluded={p.toggleSurveyExcluded}
        onGoTransfer={p.onGoTransfer}
        onToggleDisabled={p.toggleSurveyFileDisabled}
      />
    </div>
  );
}
// 반별 QR 안내문 (한 반에 A4 한 장)
function SurveyQrPrintPages({ sheet }) {
  return (
    <>
      {sheet.items.map((it, i) => (
        <div
          key={it.c}
          style={{
            pageBreakAfter: i < sheet.items.length - 1 ? "always" : "auto",
            breakAfter: i < sheet.items.length - 1 ? "page" : "auto",
            fontFamily: FONT,
            color: "#000",
            padding: "8mm 10mm",
            textAlign: "center",
          }}
        >
          <div
            style={{
              fontSize: 16,
              color: "#333",
            }}
          >
            {sheet.title}
          </div>
          <div
            style={{
              fontSize: 46,
              fontWeight: 800,
              margin: "5mm 0 5mm",
            }}
          >
            {sheet.grade}
            {"\uD559\uB144 "}
            {it.c}반
          </div>
          <div
            style={{
              display: "inline-block",
              border: "1px solid #999",
              padding: "4mm",
              background: "#fff",
            }}
          >
            <QrSvg text={it.url} size={280} />
          </div>
          <div
            style={{
              fontSize: 21,
              fontWeight: 800,
              marginTop: "5mm",
            }}
          >
            휴대폰 카메라로 QR 코드를 찍어 들어가세요
          </div>
          {sheet.deadline && (
            <div
              style={{
                fontSize: 18,
                marginTop: "3mm",
              }}
            >
              {"\uB9C8\uAC10: "}
              {sheet.deadline}
            </div>
          )}
          <div
            style={{
              textAlign: "left",
              maxWidth: "160mm",
              margin: "8mm auto 0",
              fontSize: 15.5,
              lineHeight: 1.85,
            }}
          >
            <div>
              {"\u2460 "}
              <b>우리 반</b>
              {"\uACFC "}
              <b>내 번호</b>
              {"\uB97C \uB204\uB974\uACE0 "}
              <b>이름</b>을 적어요.
            </div>
            <div>
              {"\u2461 \uD559\uAE30\uB9C8\uB2E4 "}
              <b>들었거나 신청한 과목</b>을 골라요. ‘택4’라면 꼭 4개를 골라야 해요.
            </div>
            <div>③ 전학 온 학생은 전 학교에 다닌 학기에 ‘다른 학교에 다녔어요’를 체크해요.</div>
            <div>
              {"\u2463 \uB9C8\uC9C0\uB9C9 \uD654\uBA74\uC5D0\uC11C \uD655\uC778\uD558\uACE0 "}
              <b>[제출하기]</b>를 눌러요. 잘못 냈으면 다시 내면 마지막에 낸 것으로 확인해요.
            </div>
          </div>
          <div
            style={{
              fontSize: 12,
              color: "#444",
              marginTop: "9mm",
              lineHeight: 1.6,
            }}
          >
            개인정보 안내: 적은 반·번호·이름과 고른 과목은 졸업 이수 학점 확인에만 쓰고, 확인이 끝나면 지웁니다.
            <br />
            QR 코드가 찍히지 않으면 아래 주소를 인터넷 주소창에 직접 입력하세요.
          </div>
          <div
            style={{
              fontSize: 10.5,
              color: "#444",
              wordBreak: "break-all",
              marginTop: "2mm",
            }}
          >
            {it.url}
          </div>
        </div>
      ))}
    </>
  );
}
// ================= PAGE 2 · 학기별 자료 출처 (수강신청 파일 + 기초조사 함께 쓰기) =================
const SURVEY_SOURCE_META = {
  badge: "조사",
  short: "기초조사",
  color: "#B7791F",
  title: "학생 기초조사 응답",
};
function SourceChip({ id, regFiles }) {
  if (id === SURVEY_SOURCE_ID)
    return (
      <span
        style={{
          fontSize: 11,
          fontWeight: 700,
          color: "#fff",
          background: SURVEY_SOURCE_META.color,
          borderRadius: 4,
          padding: "1px 6px",
          flexShrink: 0,
          whiteSpace: "nowrap",
        }}
      >
        {SURVEY_SOURCE_META.badge} {SURVEY_SOURCE_META.short}
      </span>
    );
  const f = regFiles.find((x) => x.id === id);
  return f ? <SourceTag source={f.source} /> : null;
}
function sourceLabelFor(id, sem, regFiles, regFileInfos, surveyData) {
  var _a;
  if (id === SURVEY_SOURCE_ID)
    return {
      label: SURVEY_SOURCE_META.title,
      students: surveyData.semesterCounts.get(sem) || 0,
    };
  const f = regFiles.find((x) => x.id === id);
  const info = regFileInfos.get(id);
  const c = info ? info.semesterCounts.find((x) => x.sem === sem) : null;
  return {
    label: f ? ((_a = f.label) !== null && _a !== void 0 ? _a : f.name) : "(지워진 파일)",
    students: c ? c.students : 0,
  };
}
function Grade1Question({ status, setAnswer, onGoCurriculum }) {
  const opt = (value, text) => (
    <label
      key={value}
      style={{
        display: "flex",
        alignItems: "center",
        gap: 6,
        fontSize: 13,
        cursor: "pointer",
        color: INK,
      }}
    >
      <input type="radio" name="grade1-electives" checked={status.answer === value} onChange={() => setAnswer(value)} />
      {text}
    </label>
  );
  const semText = status.sems.map(formatSemester).join(", ");
  let tone = MUTED;
  let msg;
  if (status.level === "ask") {
    msg = status.has
      ? `편제표에는 1학년 학생선택 과목이 ${status.courseCount}개(${semText}) 있습니다. 우리 학교와 맞는지 골라서 확인해주세요.`
      : "편제표에는 1학년 학생선택 과목이 없습니다(1학년 과목은 모두 학교지정으로 자동 반영). 우리 학교와 맞는지 골라서 확인해주세요.";
  } else if (status.level === "ok") {
    tone = OK;
    msg = status.has
      ? `✓ 편제표와 같습니다. 1학년 학생선택 과목 ${status.courseCount}개(${semText})는 아래 표의 1학년 학기에도 수강신청 파일이나 기초조사 자료를 넣어야 계산됩니다.`
      : "✓ 편제표와 같습니다. 1학년 과목은 모든 학생에게 학교지정으로 자동 반영됩니다.";
  } else {
    tone = WARN;
    msg =
      status.answer === "some"
        ? "⚠ 편제표에는 1학년 학생선택 과목이 없어서, 지금은 1학년 과목이 모두 학교지정으로 계산됩니다. ① 편제표에서 1학년 선택과목 줄의 ‘구분’을 학생선택으로, ‘선택구분’에 택N(예: 택2)을 적어 다시 올려주세요."
        : `⚠ 편제표에는 1학년 학생선택 과목이 ${status.courseCount}개(${semText}) 있습니다. 편제표가 맞다면 ‘있음’을 고르고 1학년 자료도 넣어주세요. 실제로는 모두 학교지정이라면 ① 편제표에서 그 과목의 ‘구분’을 학교지정으로 고쳐 다시 올려주세요.`;
  }
  return (
    <div
      style={{
        background: status.level === "warn" ? WARN_BG : PAPER,
        borderRadius: 8,
        padding: "10px 12px",
        marginBottom: 12,
      }}
    >
      <div
        style={{
          fontSize: 13,
          fontWeight: 800,
          marginBottom: 6,
        }}
      >
        1학년 1·2학기에 학생이 고르는 과목(학생선택)이 있나요?
      </div>
      <div
        style={{
          display: "flex",
          gap: 18,
          flexWrap: "wrap",
          marginBottom: 6,
        }}
      >
        {opt("none", "없음 — 1학년 과목은 모두 학교지정")}
        {opt("some", "있음 — 1학년 때도 학생이 과목을 골랐음")}
      </div>
      <div
        style={{
          fontSize: 12.3,
          color: tone,
          lineHeight: 1.6,
          fontWeight: status.level === "warn" ? 700 : 400,
        }}
      >
        {msg}
        {status.level === "warn" && status.answer === "some" && (
          <button
            onClick={onGoCurriculum}
            style={buttonStyle("soft", {
              fontSize: 11.5,
              padding: "3px 9px",
              marginLeft: 8,
            })}
          >
            ① 편제표로 가기
          </button>
        )}
      </div>
    </div>
  );
}
function SemesterSourcePanel({
  semesterSummary,
  providers,
  chosenSourceBySem,
  explicitChoice,
  regFiles,
  regFileInfos,
  surveyData,
  setSemesterSource,
  grade1,
  setGrade1Answer,
  onGoCurriculum,
}) {
  return (
    <div
      style={{
        ...CARD,
        padding: 16,
        marginBottom: 18,
      }}
    >
      <div
        style={{
          fontSize: 15,
          fontWeight: 800,
          marginBottom: 4,
        }}
      >
        학기별 자료 출처
      </div>
      <div
        style={{
          fontSize: 12.5,
          color: MUTED,
          lineHeight: 1.6,
          marginBottom: 10,
        }}
      >
        {
          "\uD559\uAE30\uB9C8\uB2E4 \uC5B4\uB290 \uC790\uB8CC\uB85C \uD559\uC0DD\uC120\uD0DD \uACFC\uBAA9\uC744 \uACC4\uC0B0\uD560\uC9C0 \uBCF4\uC5EC \uC90D\uB2C8\uB2E4. "
        }
        <b
          style={{
            color: INK,
          }}
        >
          수강신청 파일과 학생 기초조사를 함께 쓸 수 있습니다
        </b>
        {
          "(\uC608: 2\uD559\uB144\uC740 \uAE30\uCD08\uC870\uC0AC, 3\uD559\uB144\uC740 \uACE0\uAD50\uD559\uC810\uC81C \uB204\uB9AC\uC9D1 \uD30C\uC77C). \uC790\uB8CC\uB97C \uD569\uCE58\uBA74 \uC2E0\uCCAD\uD558\uC9C0 \uC54A\uC740 \uACFC\uBAA9\uC774 \uBD99\uC744 \uC218 \uC788\uC5B4\uC11C, "
        }
        <b
          style={{
            color: INK,
          }}
        >
          한 학기에는 자료 하나만
        </b>
        {" \uC501\uB2C8\uB2E4."}
      </div>
      <Grade1Question status={grade1} setAnswer={setGrade1Answer} onGoCurriculum={onGoCurriculum} />
      <div
        style={{
          overflowX: "auto",
        }}
      >
        <table
          style={{
            borderCollapse: "collapse",
            width: "100%",
            fontSize: 12.5,
            minWidth: 760,
          }}
        >
          <thead>
            <tr
              style={{
                background: PAPER,
              }}
            >
              <Th>학기</Th>
              <Th>편제표</Th>
              <Th>점검에 쓰는 자료</Th>
              <Th>신청 확인</Th>
            </tr>
          </thead>
          <tbody>
            {semesterSummary.map((s) => {
              const ids = providers.get(s.sem) || [];
              const chosen = chosenSourceBySem.get(s.sem);
              const unresolved = ids.length > 1 && !ids.includes(explicitChoice[s.sem]);
              const full = s.status === "auto" || s.status === "full";
              const statusText =
                s.status === "auto"
                  ? "필요 없음"
                  : s.status === "missing"
                    ? "아직 없음"
                    : `편제표 과목 ${s.elective.length}개 중 ${s.coveredCount}개가 신청됨`;
              return (
                <tr
                  key={s.sem}
                  style={{
                    borderTop: `1px solid ${LINE}`,
                    background: unresolved ? "#FDF6F2" : "transparent",
                  }}
                >
                  <Td
                    style={{
                      fontWeight: 700,
                    }}
                  >
                    {s.label}
                  </Td>
                  <Td
                    style={{
                      color: s.isAuto ? MUTED : INK,
                    }}
                  >
                    {s.isAuto ? "학교지정 과목만" : `학생선택 ${s.elective.length}과목`}
                  </Td>
                  <Td
                    style={{
                      whiteSpace: "normal",
                    }}
                  >
                    {s.isAuto ? (
                      <span
                        style={{
                          color: OK,
                        }}
                      >
                        자동 반영 (자료 필요 없음)
                      </span>
                    ) : ids.length === 0 ? (
                      <span
                        style={{
                          color: WARN,
                          fontWeight: 700,
                        }}
                      >
                        자료 없음 — 아래에서 수강신청 파일을 올리거나 기초조사 응답을 올려주세요
                      </span>
                    ) : (
                      <div
                        style={{
                          display: "flex",
                          flexDirection: "column",
                          gap: 4,
                        }}
                      >
                        {ids.map((id) => {
                          const { label, students } = sourceLabelFor(id, s.sem, regFiles, regFileInfos, surveyData);
                          const body = (
                            <>
                              <SourceChip id={id} regFiles={regFiles} />
                              <span
                                style={{
                                  overflow: "hidden",
                                  textOverflow: "ellipsis",
                                  whiteSpace: "nowrap",
                                  maxWidth: 300,
                                }}
                              >
                                {label}
                              </span>
                              <span
                                style={{
                                  color: MUTED,
                                  flexShrink: 0,
                                }}
                              >
                                {"\u00B7 "}
                                {students}명
                              </span>
                            </>
                          );
                          if (ids.length === 1)
                            return (
                              <span
                                key={id}
                                style={{
                                  display: "flex",
                                  alignItems: "center",
                                  gap: 6,
                                  minWidth: 0,
                                }}
                              >
                                {body}
                              </span>
                            );
                          return (
                            <label
                              key={id}
                              style={{
                                display: "flex",
                                alignItems: "center",
                                gap: 6,
                                cursor: "pointer",
                                minWidth: 0,
                              }}
                            >
                              <input
                                type="radio"
                                name={`sem-src-${s.sem}`}
                                checked={chosen === id}
                                onChange={() => setSemesterSource(s.sem, id)}
                                onClick={() => setSemesterSource(s.sem, id)}
                              />
                              {body}
                            </label>
                          );
                        })}
                        {unresolved && (
                          <span
                            style={{
                              color: WARN,
                              fontWeight: 700,
                            }}
                          >
                            ⚠ 이 학기 자료가 여러 곳에 있습니다. 쓸 자료를 눌러 확정해주세요 (지금은 나중에 올린 자료로
                            계산 중 — 그대로 쓰려면 골라져 있는 것을 한 번 누르세요).
                          </span>
                        )}
                      </div>
                    )}
                  </Td>
                  <Td
                    style={{
                      color: full ? OK : s.status === "missing" ? WARN : ACCENT,
                      whiteSpace: "normal",
                    }}
                  >
                    {full && !s.isAuto ? `✓ ${statusText}` : statusText}
                  </Td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}
// 수강신청 엑셀 / 학생 기초조사 화면 전환 카드 — 어느 쪽 자료를 쓸지는 위 표에서 학기마다 정해집니다
function RosterModePicker({ mode, setMode, uploadCount, surveyCount, uploadSems, surveySems }) {
  const card = (id, emoji, title, desc, count, sems) => {
    const on = mode === id;
    return (
      <button
        key={id}
        onClick={() => setMode(id)}
        style={{
          display: "block",
          textAlign: "left",
          background: on ? ACCENT_BG : "#fff",
          border: `${on ? 2 : 1.5}px solid ${on ? ACCENT : LINE}`,
          borderRadius: 12,
          padding: on ? "13px 15px" : "13.5px 15.5px",
          cursor: "pointer",
          minWidth: 0,
          fontFamily: FONT,
          color: INK,
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 8,
          }}
        >
          <span
            style={{
              fontSize: 15,
              fontWeight: 800,
            }}
          >
            <span
              style={{
                marginRight: 6,
              }}
            >
              {emoji}
            </span>
            {title}
          </span>
          {on && (
            <span
              style={{
                fontSize: 11,
                fontWeight: 700,
                color: "#fff",
                background: ACCENT,
                borderRadius: 999,
                padding: "2px 8px",
                whiteSpace: "nowrap",
              }}
            >
              지금 보는 화면
            </span>
          )}
        </div>
        <div
          style={{
            fontSize: 12.5,
            color: MUTED,
            marginTop: 4,
            lineHeight: 1.55,
          }}
        >
          {desc}
        </div>
        <div
          style={{
            fontSize: 12,
            marginTop: 6,
            fontWeight: 700,
            color: sems.length ? OK : MUTED,
          }}
        >
          {count > 0 ? `${id === "upload" ? "올린 파일" : "응답 파일"} ${count}개 · ` : ""}
          {sems.length ? `점검에 쓰는 학기: ${sems.map(formatSemester).join(", ")}` : "점검에 쓰는 학기 없음"}
        </div>
      </button>
    );
  };
  return (
    <div
      style={{
        marginBottom: 16,
      }}
    >
      <div
        style={{
          fontSize: 13.5,
          color: MUTED,
          lineHeight: 1.6,
          marginBottom: 10,
        }}
      >
        {
          "\uD559\uC0DD\uBCC4 \uC120\uD0DD\uACFC\uBAA9 \uC790\uB8CC\uB97C \uB123\uB294 \uBC29\uBC95\uC740 \uB450 \uAC00\uC9C0\uC774\uACE0 "
        }
        <b
          style={{
            color: INK,
          }}
        >
          함께 써도 됩니다
        </b>
        . 카드를 누르면 그 방법의 화면이 아래에 나옵니다.
      </div>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))",
          gap: 12,
        }}
      >
        {card(
          "upload",
          "📥",
          "수강신청 엑셀 업로드",
          "고교학점제 누리집·압핀에서 내려받은 결과 파일을 그대로 올립니다. 파일이 정확하다면 가장 빠릅니다.",
          uploadCount,
          uploadSems,
        )}
        {card(
          "survey",
          "📝",
          "학생 기초조사",
          "학생이 휴대폰으로 QR 코드를 찍어 반·번호를 누르고 자기가 들은 과목을 고릅니다. 누가 안 냈는지도 반별로 확인합니다.",
          surveyCount,
          surveySems,
        )}
      </div>
    </div>
  );
}
// 결번 입력칸: 쉼표를 치는 중에도 글자가 사라지지 않도록 입력 중인 글자는 따로 들고 있다가 번호 목록으로 바꿔 저장합니다
function SkipInput({ value, onChange, max }) {
  const [text, setText] = useState((value || []).join(", "));
  useEffect(() => {
    const cur = parseSkipNumbers(text).filter((n) => !max || n <= max);
    if (cur.join(",") !== (value || []).join(",")) setText((value || []).join(", "));
  }, [value]);
  const outOfRange = max ? parseSkipNumbers(text).filter((n) => n > max) : [];
  return (
    <span
      style={{
        display: "inline-flex",
        flexDirection: "column",
        minWidth: 0,
      }}
    >
      <input
        value={text}
        onChange={(e) => {
          setText(e.target.value);
          onChange(parseSkipNumbers(e.target.value));
        }}
        placeholder="결번 없음"
        style={{
          ...inputStyle,
          width: "100%",
          fontSize: 12.5,
          borderColor: outOfRange.length ? WARN : LINE,
        }}
      />
      {outOfRange.length > 0 && (
        <span
          style={{
            fontSize: 11,
            color: WARN,
          }}
        >
          인원({max}
          {"\uBA85)\uBCF4\uB2E4 \uD070 \uBC88\uD638: "}
          {outOfRange.join(", ")}
        </span>
      )}
    </span>
  );
}
// ================= ③ · ④ 수강신청 시스템 반영 표시 · 확인 필요 학생 목록 =================
function isChangeUnapplied(r) {
  return r.change.applyState === "pending" || r.change.applyState === "stale";
}
function classLabelText(label) {
  return label === UNCLASSIFIED_LABEL ? "반 정보 없음" : label;
}
// 상담 후 바꾼 과목을 실제 수강신청 시스템에도 고쳤는지 표시하는 체크 (④ 편집 화면·③ 목록·상세 창에서 같은 것을 씁니다)
function ChangeApplyBox({ r, setChangeApplied, compact }) {
  const st = r.change.applyState;
  if (!st || st === "none") return null;
  const applied = st === "applied";
  const color = applied ? OK : WARN;
  const when = r.change.appliedAt ? formatDateYmd(new Date(r.change.appliedAt)) : "";
  return (
    <label
      title="수강신청 시스템(고교학점제 누리집·압핀 등)에도 이 변경을 반영했으면 체크하세요"
      style={{
        display: "flex",
        gap: 7,
        alignItems: "flex-start",
        cursor: "pointer",
        background: compact ? "transparent" : applied ? OK_BG : WARN_BG,
        borderRadius: 8,
        padding: compact ? 0 : "10px 12px",
        marginTop: compact ? 0 : 12,
      }}
    >
      <input
        type="checkbox"
        checked={applied}
        onChange={(e) => setChangeApplied(r, e.target.checked)}
        style={{
          marginTop: 2,
          flexShrink: 0,
        }}
      />
      <span
        style={{
          fontSize: compact ? 12 : 13,
          fontWeight: 700,
          color,
          lineHeight: 1.5,
        }}
      >
        {compact
          ? applied
            ? `시스템 반영함 (${when})`
            : st === "stale"
              ? "반영 후 다시 바뀜 — 재반영 필요"
              : "시스템 미반영"
          : "수강신청 시스템(고교학점제 누리집·압핀 등)에도 이 변경을 반영했습니다"}
        {!compact && (
          <span
            style={{
              display: "block",
              fontWeight: 400,
              fontSize: 12,
              color,
              marginTop: 2,
            }}
          >
            {applied
              ? `✓ ${when}에 반영 표시함. ③ 점검 화면·확인서·CSV에 ‘반영함’으로 나옵니다.`
              : st === "stale"
                ? "⚠ 반영 표시를 한 뒤 변경 내역이 또 바뀌었습니다. 바뀐 내용을 수강신청 시스템에 다시 반영한 뒤 체크하세요."
                : "수강신청 시스템에서 과목을 똑같이 고친 뒤 체크하세요. 체크하기 전까지 ③ 화면에 ‘변경·미반영’으로 표시됩니다."}
          </span>
        )}
      </span>
    </label>
  );
}
// ③ 요약 카드에서 ‘확인 필요’ 또는 ‘변경 미반영’을 누르면 나오는 학생 목록
function ProblemListPanel({
  mode,
  rows,
  selectedClass,
  onClearClass,
  onClose,
  onOpenDetail,
  onOpenChange,
  printOne,
  printFailOnly,
  setChangeApplied,
  onPrintList,
  panelRef,
}) {
  const [problemKey, setProblemKey] = useState(null);
  useEffect(() => setProblemKey(null), [mode]);
  const isFail = mode === "fail";
  const categories = useMemo(() => {
    const m = new Map();
    rows.forEach((r) => {
      const seen = new Set();
      r.reasons.forEach((x) => {
        const cat = reasonCategory(x);
        if (seen.has(cat.key)) return;
        seen.add(cat.key);
        if (!m.has(cat.key))
          m.set(cat.key, {
            ...cat,
            count: 0,
          });
        m.get(cat.key).count += 1;
      });
    });
    return Array.from(m.values()).sort((a, b) => (a.blocking === b.blocking ? b.count - a.count : a.blocking ? -1 : 1));
  }, [rows]);
  const activeCat = isFail && problemKey ? categories.find((c) => c.key === problemKey) : null;
  const shown = (activeCat ? rows.filter((r) => r.reasons.some((x) => reasonCategory(x).key === activeCat.key)) : rows)
    .slice()
    .sort(byClassThenId);
  const byClass = [];
  shown.forEach((r) => {
    const last = byClass[byClass.length - 1];
    if (last && last.label === r.classLabel) last.rows.push(r);
    else
      byClass.push({
        label: r.classLabel,
        rows: [r],
      });
  });
  const semList = (list) => list.map((e) => `${e.name}(${formatSemester(e.semester)})`).join(", ");
  const tone = isFail ? WARN : "#B7791F";
  const scope = [selectedClass ? classLabelText(selectedClass) : "전체 학급", activeCat ? activeCat.label : null]
    .filter(Boolean)
    .join(" · ");
  const printThis = () =>
    onPrintList({
      title: isFail ? "확인 필요 학생 목록" : "수강신청 시스템 반영이 필요한 학생 목록",
      subtitle: `${scope} · ${shown.length}명`,
      columns: isFail
        ? ["학급", "학번", "이름", "판정 원인 (확인필요가 된 이유)", "참고"]
        : ["학급", "학번", "이름", "추가한 과목", "취소한 과목", "시스템 반영 확인"],
      rows: shown.map((r) =>
        isFail
          ? [
              classLabelText(r.classLabel),
              r.studentId,
              r.name,
              r.reasons
                .filter((x) => x.blocking)
                .map((x) => x.text)
                .join(", ") || "-",
              r.reasons
                .filter((x) => !x.blocking)
                .map((x) => x.text)
                .join(", "),
            ]
          : [
              classLabelText(r.classLabel),
              r.studentId,
              r.name,
              semList(r.change.added) || "-",
              semList(r.change.removed) || "-",
              "",
            ],
      ),
    });
  const small = (kind) =>
    buttonStyle(kind, {
      fontSize: 11.5,
      padding: "4px 9px",
    });
  return (
    <div
      ref={panelRef}
      style={{
        ...CARD,
        padding: 16,
        marginBottom: 16,
        borderTop: `3px solid ${tone}`,
        scrollMarginTop: 120,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          gap: 10,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            minWidth: 0,
          }}
        >
          <div
            style={{
              fontSize: 15.5,
              fontWeight: 800,
              color: INK,
            }}
          >
            {isFail ? "확인 필요 학생" : "변경했지만 수강신청 시스템 반영 표시가 없는 학생"} {shown.length}명
            <span
              style={{
                fontSize: 12.5,
                fontWeight: 600,
                color: MUTED,
                marginLeft: 8,
              }}
            >
              {scope}
            </span>
            {selectedClass && (
              <button
                onClick={onClearClass}
                style={{
                  ...chipStyle(true),
                  marginLeft: 6,
                }}
              >
                학급 선택 풀기 ✕
              </button>
            )}
          </div>
          <div
            style={{
              fontSize: 12,
              color: MUTED,
              marginTop: 3,
              lineHeight: 1.55,
            }}
          >
            {isFail
              ? "빨간 항목이 ‘확인필요’가 된 원인이고, 회색 항목은 참고용입니다. 아래 문제를 누르면 그 문제가 있는 학생만 보입니다."
              : "④에서 과목을 바꾼 뒤 아직 ‘수강신청 시스템에 반영했습니다’를 체크하지 않은 학생입니다. 시스템에서 고친 뒤 체크하면 이 목록에서 빠집니다."}
          </div>
        </div>
        <div
          style={{
            display: "flex",
            gap: 6,
            flexWrap: "wrap",
          }}
        >
          <button
            onClick={printThis}
            disabled={shown.length === 0}
            style={buttonStyle(shown.length ? "soft" : "disabled")}
          >
            🖨 이 목록 인쇄
          </button>
          {isFail && (
            <button onClick={printFailOnly} style={buttonStyle("ghost")}>
              🖨 확인서 인쇄 (확인필요 학생)
            </button>
          )}
          <button onClick={onClose} style={buttonStyle("ghost")}>
            닫기
          </button>
        </div>
      </div>
      {isFail && categories.length > 0 && (
        <div
          style={{
            display: "flex",
            gap: 6,
            flexWrap: "wrap",
            alignItems: "center",
            marginTop: 10,
          }}
        >
          <button onClick={() => setProblemKey(null)} style={chipStyle(!problemKey)}>
            {"\uBAA8\uB4E0 \uBB38\uC81C "}
            {rows.length}
          </button>
          {categories.map((c) => (
            <button
              key={c.key}
              onClick={() => setProblemKey(problemKey === c.key ? null : c.key)}
              style={{
                ...chipStyle(problemKey === c.key),
                color: problemKey === c.key ? ACCENT : c.blocking ? WARN : MUTED,
                borderColor: problemKey === c.key ? ACCENT : c.blocking ? "#E8C9BB" : LINE,
                background: problemKey === c.key ? ACCENT_BG : c.blocking ? WARN_BG : "#fff",
              }}
            >
              {c.label} {c.count}
            </button>
          ))}
        </div>
      )}
      <div
        style={{
          overflowX: "auto",
          marginTop: 12,
        }}
      >
        <table
          style={{
            borderCollapse: "collapse",
            width: "100%",
            fontSize: 12.5,
            minWidth: 860,
          }}
        >
          <thead>
            <tr
              style={{
                background: PAPER,
              }}
            >
              <Th>학번</Th>
              <Th>이름</Th>
              <Th>{isFail ? "확인 사유" : "변경 내역"}</Th>
              <Th>{isFail ? "상담·변경" : "수강신청 시스템 반영"}</Th>
              <Th>바로가기</Th>
            </tr>
          </thead>
          <tbody>
            {shown.length === 0 && (
              <tr>
                <td
                  colSpan={5}
                  style={{
                    padding: 18,
                    textAlign: "center",
                    color: MUTED,
                  }}
                >
                  {isFail ? "확인 필요 학생이 없습니다." : "반영 표시가 필요한 학생이 없습니다."}
                </td>
              </tr>
            )}
            {byClass.map((g) => (
              <React.Fragment key={g.label}>
                <tr
                  style={{
                    background: "#FBFAF7",
                    borderTop: `1px solid ${LINE}`,
                  }}
                >
                  <td
                    colSpan={5}
                    style={{
                      padding: "6px 12px",
                      fontSize: 12,
                      fontWeight: 800,
                      color: g.label === UNCLASSIFIED_LABEL ? WARN : INK,
                    }}
                  >
                    {classLabelText(g.label)}
                    {" \u00B7 "}
                    {g.rows.length}명
                  </td>
                </tr>
                {g.rows.map((r) => (
                  <tr
                    key={r.key}
                    style={{
                      borderTop: `1px solid ${LINE}`,
                      verticalAlign: "top",
                    }}
                  >
                    <Td
                      style={{
                        color: MUTED,
                      }}
                    >
                      {r.studentId}
                    </Td>
                    <Td
                      style={{
                        fontWeight: 700,
                      }}
                    >
                      {r.name}
                      <StudentTags r={r} />
                    </Td>
                    <Td
                      style={{
                        whiteSpace: "normal",
                        minWidth: 280,
                      }}
                    >
                      {isFail ? (
                        <ReasonList reasons={r.reasons} compact={true} />
                      ) : (
                        <div
                          style={{
                            lineHeight: 1.6,
                          }}
                        >
                          {r.change.added.map((e, i) => (
                            <div
                              key={`a${i}`}
                              style={{
                                color: OK,
                              }}
                            >
                              {"+ \uCD94\uAC00 \u00B7 "}
                              {e.name}
                              {" ("}
                              {formatSemester(e.semester)})
                            </div>
                          ))}
                          {r.change.removed.map((e, i) => (
                            <div
                              key={`r${i}`}
                              style={{
                                color: WARN,
                              }}
                            >
                              {"\u2212 \uCDE8\uC18C \u00B7 "}
                              {e.name}
                              {" ("}
                              {formatSemester(e.semester)})
                            </div>
                          ))}
                        </div>
                      )}
                    </Td>
                    <Td
                      style={{
                        whiteSpace: "normal",
                        minWidth: 150,
                      }}
                    >
                      {r.change.applyState !== "none" ? (
                        <ChangeApplyBox r={r} setChangeApplied={setChangeApplied} compact={true} />
                      ) : (
                        <span
                          style={{
                            fontSize: 12,
                            color: MUTED,
                          }}
                        >
                          {r.hasChanges ? "상담 메모만 있음" : "아직 없음"}
                        </span>
                      )}
                    </Td>
                    <Td>
                      <span
                        style={{
                          display: "flex",
                          gap: 4,
                        }}
                      >
                        <button onClick={() => onOpenDetail(r.key)} style={small("ghost")}>
                          상세
                        </button>
                        <button onClick={() => onOpenChange(r.key)} style={small("soft")}>
                          ④ 변경
                        </button>
                        <button onClick={() => printOne(r.key)} style={small("ghost")}>
                          🖨
                        </button>
                      </span>
                    </Td>
                  </tr>
                ))}
              </React.Fragment>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}
// ‘이 목록 인쇄’ — 담임·업무 담당 선생님께 나눠 줄 A4 표
function ListPrintPage({ sheet }) {
  return (
    <div
      style={{
        fontFamily: FONT,
        color: "#000",
        padding: "4mm 2mm",
      }}
    >
      <div
        style={{
          fontSize: 18,
          fontWeight: 800,
        }}
      >
        {sheet.title}
      </div>
      <div
        style={{
          fontSize: 12,
          color: "#444",
          margin: "2mm 0 4mm",
        }}
      >
        {sheet.subtitle}
        {" \u00B7 \uCD9C\uB825 "}
        {formatDateYmd(new Date())}
      </div>
      <table
        style={{
          borderCollapse: "collapse",
          width: "100%",
          fontSize: 11,
        }}
      >
        <thead>
          <tr>
            {sheet.columns.map((c) => (
              <th
                key={c}
                style={{
                  border: "1px solid #999",
                  padding: "3px 5px",
                  background: "#eee",
                  textAlign: "left",
                  whiteSpace: "nowrap",
                }}
              >
                {c}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {sheet.rows.map((row, i) => (
            <tr
              key={i}
              style={{
                breakInside: "avoid",
                pageBreakInside: "avoid",
              }}
            >
              {row.map((cell, j) => (
                <td
                  key={j}
                  style={{
                    border: "1px solid #999",
                    padding: "3px 5px",
                    verticalAlign: "top",
                    minWidth: sheet.columns[j] === "시스템 반영 확인" ? 90 : undefined,
                  }}
                >
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      <div
        style={{
          fontSize: 10,
          color: "#555",
          marginTop: "4mm",
        }}
      >
        학생 개인정보가 담긴 문서입니다. 업무에만 사용하고 사용이 끝나면 파기하세요.
      </div>
    </div>
  );
}
// ================= PAGE 3: 졸업요건 점검 =================
const CARD = {
  background: "#fff",
  border: `1px solid ${LINE}`,
  borderRadius: 10,
};
const inputStyle = {
  fontSize: 12.5,
  padding: "5px 7px",
  borderRadius: 6,
  border: `1px solid ${LINE}`,
  background: "#fff",
  color: INK,
  fontFamily: FONT,
  boxSizing: "border-box",
};
const SPECIALTY_GROUPS = ["과학계열", "체육계열", "예술계열", "외국어·국제계열"];
function buttonStyle(kind, extra) {
  const base = {
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
    fontSize: 12.5,
    fontWeight: 700,
    borderRadius: 7,
    padding: "6px 11px",
    cursor: "pointer",
    whiteSpace: "nowrap",
    fontFamily: FONT,
  };
  const kinds = {
    primary: {
      color: "#fff",
      background: ACCENT,
      border: "none",
    },
    soft: {
      color: ACCENT,
      background: ACCENT_BG,
      border: "none",
    },
    ghost: {
      color: INK,
      background: "#fff",
      border: `1px solid ${LINE}`,
    },
    danger: {
      color: WARN,
      background: "#fff",
      border: "1px solid #E8C9BB",
    },
    disabled: {
      color: "#fff",
      background: "#B8B4A9",
      border: "none",
      cursor: "default",
    },
  };
  return {
    ...base,
    ...(kinds[kind] || kinds.soft),
    ...(extra || {}),
  };
}
function chipStyle(on) {
  return {
    fontSize: 11.5,
    fontWeight: 700,
    padding: "3px 9px",
    borderRadius: 999,
    border: `1px solid ${on ? ACCENT : LINE}`,
    background: on ? ACCENT_BG : "#fff",
    color: on ? ACCENT : MUTED,
    cursor: "pointer",
    fontFamily: FONT,
    whiteSpace: "nowrap",
  };
}
function PassBadge({ pass, size = 11.5 }) {
  return (
    <span
      style={{
        fontSize: size,
        fontWeight: 700,
        padding: "3px 9px",
        borderRadius: 999,
        background: pass ? OK_BG : WARN_BG,
        color: pass ? OK : WARN,
        whiteSpace: "nowrap",
      }}
    >
      {pass ? "충족" : "확인필요"}
    </span>
  );
}
function StudentTags({ r }) {
  const tag = (text, color, bg) => (
    <span
      key={text}
      style={{
        marginLeft: 5,
        fontSize: 10.5,
        fontWeight: 700,
        color,
        background: bg,
        borderRadius: 999,
        padding: "1px 6px",
        whiteSpace: "nowrap",
      }}
    >
      {text}
    </span>
  );
  return (
    <>
      {r.isUnclassified && tag("반 정보 없음", MUTED, PAPER)}
      {r.isManual && tag("직접 추가", MUTED, PAPER)}
      {r.isTransfer && tag("전입", "#6B4C9A", "#EFE8F6")}
      {r.hasChanges &&
        (r.change.applyState === "applied"
          ? tag("변경·반영", OK, OK_BG)
          : r.change.applyState === "stale"
            ? tag("변경·재반영 필요", WARN, WARN_BG)
            : r.change.applyState === "pending"
              ? tag("변경·미반영", WARN, WARN_BG)
              : tag("상담", OK, OK_BG))}
      {r.extraEntries.length > 0 && tag("공동", ACCENT, ACCENT_BG)}
      {r.missingSemesters.length > 0 && tag("신청 누락", WARN, WARN_BG)}
    </>
  );
}
function ReasonList({ reasons, compact }) {
  if (!reasons || reasons.length === 0) return null;
  return (
    <div
      style={{
        display: "flex",
        flexWrap: "wrap",
        gap: 4,
      }}
    >
      {reasons.map((x, i) => (
        <span
          key={i}
          style={{
            fontSize: compact ? 11 : 12,
            fontWeight: x.blocking ? 700 : 500,
            color: x.blocking ? WARN : MUTED,
            background: x.blocking ? WARN_BG : PAPER,
            borderRadius: 6,
            padding: compact ? "1px 6px" : "3px 8px",
            whiteSpace: "nowrap",
          }}
        >
          {x.text}
        </span>
      ))}
    </div>
  );
}
function GroupCreditGrid({ result, groups, reqOverride, priorityGroups }) {
  return (
    <div
      style={{
        display: "grid",
        gridTemplateColumns: "repeat(auto-fill, minmax(118px, 1fr))",
        gap: 6,
      }}
    >
      {groups.map((g) => {
        const cur = result.groupCredit[g] || 0;
        const req = Number(reqOverride[g]) || 0;
        const short = cur < req;
        return (
          <div
            key={g}
            style={{
              background: short
                ? WARN_BG
                : (priorityGroups === null || priorityGroups === void 0 ? void 0 : priorityGroups[g])
                  ? "#FDF3C7"
                  : PAPER,
              borderRadius: 7,
              padding: "6px 9px",
            }}
          >
            <div
              style={{
                fontSize: 11.5,
                color: short ? WARN : MUTED,
                fontWeight: 600,
              }}
            >
              {g}
            </div>
            <div
              style={{
                fontSize: 14,
                fontWeight: 700,
                color: short ? WARN : INK,
              }}
            >
              {cur}
              {" / "}
              {req}
              {short && (
                <span
                  style={{
                    fontSize: 11,
                  }}
                >
                  {" ("}
                  {Math.round((req - cur) * 100) / 100}
                  {" \uBD80\uC871)"}
                </span>
              )}
            </div>
          </div>
        );
      })}
      {(result.otherGroups || []).map((o) => (
        <div
          key={o.group}
          style={{
            background: PAPER,
            borderRadius: 7,
            padding: "6px 9px",
          }}
        >
          <div
            style={{
              fontSize: 11.5,
              color: MUTED,
              fontWeight: 600,
            }}
          >
            {o.group}
          </div>
          <div
            style={{
              fontSize: 14,
              fontWeight: 700,
            }}
          >
            {o.credit}
          </div>
        </div>
      ))}
    </div>
  );
}
function PageIntro({ title, children }) {
  return (
    <div
      style={{
        marginBottom: 16,
      }}
    >
      <div
        style={{
          fontSize: 16,
          fontWeight: 800,
          marginBottom: 4,
        }}
      >
        {title}
      </div>
      <div
        style={{
          fontSize: 13,
          color: MUTED,
          lineHeight: 1.65,
        }}
      >
        {children}
      </div>
    </div>
  );
}
function EmptyPanel({ text }) {
  return (
    <div
      style={{
        ...CARD,
        padding: "36px 20px",
        textAlign: "center",
        color: MUTED,
        fontSize: 13.5,
        lineHeight: 1.6,
      }}
    >
      {text}
    </div>
  );
}
function DetailSection({ title, action, children }) {
  return (
    <div
      style={{
        marginTop: 20,
        borderTop: `1px solid ${LINE}`,
        paddingTop: 14,
      }}
    >
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: 8,
          gap: 8,
        }}
      >
        <div
          style={{
            fontSize: 13,
            fontWeight: 700,
          }}
        >
          {title}
        </div>
        {action}
      </div>
      {children}
    </div>
  );
}
function CurriculumIssues({ issues }) {
  if (!issues || issues.length === 0) {
    return (
      <div
        style={{
          marginTop: 16,
          background: OK_BG,
          color: OK,
          borderRadius: 10,
          padding: "10px 14px",
          fontSize: 12.5,
          fontWeight: 600,
        }}
      >
        ✓ 편제표 점검: 학점·개설학기·구분·택N 표기에 빈칸이나 형식 오류가 없습니다.
      </div>
    );
  }
  return (
    <div
      style={{
        ...CARD,
        marginTop: 16,
        padding: 14,
      }}
    >
      <div
        style={{
          fontSize: 13.5,
          fontWeight: 700,
          marginBottom: 6,
        }}
      >
        편제표 점검 — 확인이 필요한 항목
      </div>
      {issues.map((it, i) => (
        <details
          key={i}
          style={{
            marginBottom: 6,
          }}
          open={it.level === "error"}
        >
          <summary
            style={{
              cursor: "pointer",
              fontSize: 12.5,
              fontWeight: 700,
              color: it.level === "error" ? WARN : ACCENT,
            }}
          >
            {it.level === "error" ? "⚠" : "ℹ"} {it.title}
            {" ("}
            {it.items.length})
          </summary>
          <div
            style={{
              fontSize: 12,
              color: MUTED,
              lineHeight: 1.6,
              padding: "4px 0 0 18px",
            }}
          >
            {it.items.slice(0, 30).join(" · ")}
            {it.items.length > 30 && ` 외 ${it.items.length - 30}개`}
          </div>
        </details>
      ))}
    </div>
  );
}
function OrphanPanel({ items, results, onRelink, onDiscard }) {
  const [choice, setChoice] = useState({});
  if (!items || items.length === 0) return null;
  const sorted = [...results].sort((a, b) => a.name.localeCompare(b.name, "ko"));
  return (
    <div
      style={{
        background: WARN_BG,
        border: "1px solid #E8C9BB",
        borderRadius: 10,
        padding: 14,
        marginBottom: 16,
      }}
    >
      <div
        style={{
          fontSize: 13.5,
          fontWeight: 700,
          color: WARN,
        }}
      >
        {"\uD559\uC0DD\uACFC \uC5F0\uACB0\uC774 \uB04A\uAE34 \uC785\uB825 "}
        {items.length}건 (지금은 계산에 들어가지 않습니다)
      </div>
      <div
        style={{
          fontSize: 12.5,
          color: WARN,
          margin: "4px 0 10px",
          lineHeight: 1.55,
        }}
      >
        수강신청 파일을 다시 올리거나 동명이인 병합을 바꾸면, 전에 입력한 내용이 가리키던 학생을 찾지 못할 수 있습니다.
        같은 학생을 골라 다시 연결하거나, 필요 없으면 삭제하세요.
      </div>
      {items.map((it) => {
        var _a;
        const same = sorted.filter((r) => r.name === it.name);
        const value = (_a = choice[it.id]) !== null && _a !== void 0 ? _a : same.length === 1 ? same[0].key : "";
        return (
          <div
            key={it.id}
            style={{
              display: "flex",
              gap: 8,
              alignItems: "center",
              flexWrap: "wrap",
              fontSize: 12.5,
              marginBottom: 6,
            }}
          >
            <b
              style={{
                color: INK,
                minWidth: 60,
              }}
            >
              {it.name || "(이름 모름)"}
            </b>
            <span
              style={{
                color: MUTED,
              }}
            >
              {it.label}
            </span>
            <select
              value={value}
              onChange={(e) =>
                setChoice((p) => ({
                  ...p,
                  [it.id]: e.target.value,
                }))
              }
              style={{
                ...inputStyle,
                maxWidth: 280,
              }}
            >
              <option value="">연결할 학생 선택</option>
              {same.length > 0 && (
                <optgroup label="같은 이름">
                  {same.map((r) => (
                    <option key={r.key} value={r.key}>
                      {r.name}
                      {" \u00B7 "}
                      {r.studentId}
                      {" \u00B7 "}
                      {r.classLabel}
                    </option>
                  ))}
                </optgroup>
              )}
              <optgroup label="전체 학생">
                {sorted
                  .filter((r) => r.name !== it.name)
                  .map((r) => (
                    <option key={r.key} value={r.key}>
                      {r.name}
                      {" \u00B7 "}
                      {r.studentId}
                      {" \u00B7 "}
                      {r.classLabel}
                    </option>
                  ))}
              </optgroup>
            </select>
            <button
              disabled={!value}
              onClick={() => value && onRelink(it, value)}
              style={buttonStyle(value ? "primary" : "disabled", {
                fontSize: 12,
                padding: "4px 10px",
              })}
            >
              다시 연결
            </button>
            <button
              onClick={() =>
                window.confirm(`${it.name || "이 학생"}의 ‘${it.label}’ 입력을 삭제할까요?`) && onDiscard(it)
              }
              style={buttonStyle("danger", {
                fontSize: 12,
                padding: "4px 10px",
              })}
            >
              삭제
            </button>
          </div>
        );
      })}
    </div>
  );
}
function GroupSelect({ value, groups, onChange, style }) {
  const extra = value && !groups.includes(value) && !SPECIALTY_GROUPS.includes(value) ? [value] : [];
  return (
    <select
      value={value || ""}
      onChange={(e) => onChange(e.target.value)}
      style={{
        ...inputStyle,
        ...(style || {}),
      }}
    >
      <option value="">교과(군) 선택</option>
      {groups.map((g) => (
        <option key={g} value={g}>
          {g}
        </option>
      ))}
      {extra.map((g) => (
        <option key={g} value={g}>
          {g}
        </option>
      ))}
      <optgroup label="전문 교과 (총학점에만 반영)">
        {SPECIALTY_GROUPS.map((g) => (
          <option key={g} value={g}>
            {g}
          </option>
        ))}
      </optgroup>
    </select>
  );
}
function SubjectDatalist({ id, lookup }) {
  const names = useMemo(
    () => Array.from(new Set(Array.from(lookup.values()).map((s) => s.name))).sort((a, b) => a.localeCompare(b, "ko")),
    [lookup],
  );
  return (
    <datalist id={id}>
      {names.map((n) => (
        <option key={n} value={n} />
      ))}
    </datalist>
  );
}
function classOptionsOf(results) {
  const m = new Map();
  results.forEach((r) => {
    if (!m.has(r.classLabel)) m.set(r.classLabel, r.classSortKey);
  });
  return Array.from(m.entries())
    .sort((a, b) => a[1] - b[1] || a[0].localeCompare(b[0], "ko"))
    .map(([label]) => label);
}
function byClassThenId(a, b) {
  return (
    a.classSortKey - b.classSortKey ||
    a.studentId.localeCompare(b.studentId, "ko", {
      numeric: true,
    }) ||
    a.name.localeCompare(b.name, "ko")
  );
}
function StudentListPanel({ title, results, filters, filterId, onFilter, selectedKey, onSelect, emptyText }) {
  const [search, setSearch] = useState("");
  const [cls, setCls] = useState("");
  const classOptions = useMemo(() => classOptionsOf(results), [results]);
  const active = filters.find((f) => f.id === filterId) || filters[0];
  const q = search.trim();
  const list = results
    .filter(active.test)
    .filter((r) => !cls || r.classLabel === cls)
    .filter((r) => !q || r.name.includes(q) || r.studentId.includes(q))
    .sort(byClassThenId);
  return (
    <div
      style={{
        ...CARD,
        padding: 12,
        position: "sticky",
        top: 104,
        maxHeight: "calc(100vh - 124px)",
        display: "flex",
        flexDirection: "column",
        minWidth: 0,
      }}
    >
      {title && (
        <div
          style={{
            fontSize: 13,
            fontWeight: 700,
            marginBottom: 8,
          }}
        >
          {title}
        </div>
      )}
      <div
        style={{
          display: "flex",
          gap: 4,
          flexWrap: "wrap",
          marginBottom: 8,
        }}
      >
        {filters.map((f) => (
          <button key={f.id} onClick={() => onFilter(f.id)} style={chipStyle(f.id === active.id)}>
            {f.label} {results.filter(f.test).length}
          </button>
        ))}
      </div>
      <input
        value={search}
        onChange={(e) => setSearch(e.target.value)}
        placeholder="이름·학번 검색"
        style={{
          ...inputStyle,
          width: "100%",
          marginBottom: 6,
        }}
      />
      <select
        value={cls}
        onChange={(e) => setCls(e.target.value)}
        style={{
          ...inputStyle,
          width: "100%",
        }}
      >
        <option value="">모든 학급</option>
        {classOptions.map((c) => (
          <option key={c} value={c}>
            {c}
          </option>
        ))}
      </select>
      <div
        style={{
          overflowY: "auto",
          flex: 1,
          marginTop: 8,
          display: "flex",
          flexDirection: "column",
          gap: 2,
        }}
      >
        {list.length === 0 && (
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              padding: 8,
              lineHeight: 1.5,
            }}
          >
            {emptyText}
          </div>
        )}
        {list.map((r) => (
          <button
            key={r.key}
            onClick={() => onSelect(r.key)}
            style={{
              textAlign: "left",
              border: "none",
              borderRadius: 7,
              padding: "7px 9px",
              cursor: "pointer",
              background: r.key === selectedKey ? ACCENT_BG : "transparent",
              fontFamily: FONT,
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 6,
              }}
            >
              <span
                style={{
                  fontSize: 13,
                  fontWeight: 700,
                  color: INK,
                }}
              >
                {r.name || "(이름 없음)"}
              </span>
              <PassBadge pass={r.pass} size={10.5} />
            </div>
            <div
              style={{
                fontSize: 11.5,
                color: MUTED,
                marginTop: 1,
              }}
            >
              {r.studentId}
              {" \u00B7 "}
              {r.classLabel}
            </div>
          </button>
        ))}
      </div>
    </div>
  );
}
function StudentAdder({ results, excludeKeys, onAdd }) {
  const [q, setQ] = useState("");
  const [cls, setCls] = useState("");
  const classOptions = useMemo(() => classOptionsOf(results), [results]);
  const query = q.trim();
  const matches =
    query || cls
      ? results
          .filter(
            (r) =>
              !excludeKeys.has(r.key) &&
              (!cls || r.classLabel === cls) &&
              (!query || r.name.includes(query) || r.studentId.includes(query)),
          )
          .sort(byClassThenId)
      : [];
  const shown = matches.slice(0, 40);
  return (
    <div
      style={{
        marginTop: 8,
      }}
    >
      <div
        style={{
          display: "flex",
          gap: 6,
          flexWrap: "wrap",
        }}
      >
        <input
          value={q}
          onChange={(e) => setQ(e.target.value)}
          placeholder="학생 추가: 이름·학번 검색"
          style={{
            ...inputStyle,
            flex: "1 1 180px",
          }}
        />
        <select
          value={cls}
          onChange={(e) => setCls(e.target.value)}
          style={{
            ...inputStyle,
            flex: "0 1 180px",
          }}
        >
          <option value="">또는 학급에서 고르기</option>
          {classOptions.map((c) => (
            <option key={c} value={c}>
              {c}
            </option>
          ))}
        </select>
        {matches.length > 1 && (
          <button
            onClick={() => matches.forEach((r) => onAdd(r))}
            style={buttonStyle("soft", {
              fontSize: 12,
            })}
          >
            {"\uC544\uB798 "}
            {matches.length}명 모두 추가
          </button>
        )}
      </div>
      {(query || cls) && (
        <div
          style={{
            display: "flex",
            flexWrap: "wrap",
            gap: 5,
            marginTop: 6,
            maxHeight: 150,
            overflowY: "auto",
          }}
        >
          {shown.length === 0 && (
            <span
              style={{
                fontSize: 12,
                color: MUTED,
              }}
            >
              추가할 학생이 없습니다.
            </span>
          )}
          {shown.map((r) => (
            <button
              key={r.key}
              onClick={() => onAdd(r)}
              style={{
                ...chipStyle(false),
                color: INK,
              }}
            >
              {"+ "}
              {r.name}{" "}
              <span
                style={{
                  color: MUTED,
                  fontWeight: 400,
                }}
              >
                {r.studentId}
              </span>
            </button>
          ))}
          {matches.length > shown.length && (
            <span
              style={{
                fontSize: 12,
                color: MUTED,
              }}
            >
              {"\uC678 "}
              {matches.length - shown.length}명 — 검색어를 더 입력하세요
            </span>
          )}
        </div>
      )}
    </div>
  );
}
function CheckPage({
  curriculumFile,
  groups,
  priorityGroups,
  reqOverride,
  minCredit,
  results,
  allResults,
  filtered,
  passCount,
  failCount,
  search,
  setSearch,
  statusFilter,
  setStatusFilter,
  statusFilters,
  selected,
  setSelected,
  selectedResult,
  handleExport,
  classSummary,
  selectedClass,
  setSelectedClass,
  archivedResults,
  archiveModalKey,
  setArchiveModalKey,
  archiveReason,
  setArchiveReason,
  confirmArchive,
  restoreStudent,
  printOne,
  printClass,
  classTeachers,
  setClassTeachers,
  onOpenChange,
  onOpenTransfer,
  onOpenExtra,
  orphanItems,
  relinkOrphan,
  discardOrphan,
  surveyMissingByClass,
  sourceIssues,
  onGoRoster,
  setChangeApplied,
  onPrintList,
}) {
  const [showArchived, setShowArchived] = useState(false);
  const panelRef = useRef(null);
  const tableRef = useRef(null);
  const orphanPanel = (
    <OrphanPanel items={orphanItems} results={allResults} onRelink={relinkOrphan} onDiscard={discardOrphan} />
  );
  // an archived-only list must still be reachable, otherwise archiving every student would hide the restore buttons
  if (results.length === 0 && archivedResults.length === 0) {
    return (
      <div>
        {orphanPanel}
        <EmptyPanel text="편제표와 수강신청 파일을 모두 업로드하면 점검 결과가 여기에 표시됩니다." />
      </div>
    );
  }
  const missingTotal = results.filter((r) => r.missingSemesters.length > 0).length;
  const transferTotal = results.filter((r) => r.isTransfer).length;
  const changedTotal = results.filter((r) => r.hasChanges).length;
  const unappliedTotal = results.filter(isChangeUnapplied).length;
  // 요약 카드를 누르면 그 학생들로 거르고, 확인 필요·변경 미반영은 아래 목록 상자로 이동합니다 (다시 누르면 전체)
  const pickStatus = (id) => {
    const next = statusFilter === id && id !== "all" ? "all" : id;
    setStatusFilter(next);
    setTimeout(() => {
      const target = next === "fail" || next === "unapplied" ? panelRef.current : tableRef.current;
      if (target && target.scrollIntoView)
        target.scrollIntoView({
          behavior: "smooth",
          block: "start",
        });
    }, 60);
  };
  const surveyNotSubmitted = surveyMissingByClass
    ? Array.from(surveyMissingByClass.values()).reduce((a, list) => a + list.length, 0)
    : 0;
  const surveyMissingOnlyClasses = surveyMissingByClass
    ? Array.from(surveyMissingByClass.keys()).filter((label) => !classSummary.some((c) => c.label === label))
    : [];
  const small = (kind) =>
    buttonStyle(kind, {
      fontSize: 11.5,
      padding: "4px 9px",
    });
  return (
    <div>
      {orphanPanel}
      {sourceIssues && (sourceIssues.unresolvedSems.length > 0 || sourceIssues.grade1Warn) && (
        <div
          style={{
            background: WARN_BG,
            border: "1px solid #E8C9BB",
            borderRadius: 10,
            padding: "10px 14px",
            marginBottom: 14,
            fontSize: 12.8,
            color: WARN,
            lineHeight: 1.6,
            display: "flex",
            gap: 10,
            alignItems: "center",
            flexWrap: "wrap",
          }}
        >
          <span
            style={{
              flex: 1,
              minWidth: 260,
            }}
          >
            {sourceIssues.unresolvedSems.length > 0 && (
              <div>
                {"\u26A0 "}
                {sourceIssues.unresolvedSems.map(formatSemester).join(", ")}: 같은 학기 자료가 여러 곳(수강신청
                파일·기초조사)에 있어 확인이 필요합니다. 지금은 나중에 올린 자료로 계산했습니다.
              </div>
            )}
            {sourceIssues.grade1Warn && (
              <div>
                ⚠ 1학년 학생선택 과목이 있는지에 대한 답이 편제표와 다릅니다. ② 화면 ‘학기별 자료 출처’를 확인하세요.
              </div>
            )}
          </span>
          <button onClick={onGoRoster} style={buttonStyle("danger")}>
            ② 화면에서 확인
          </button>
        </div>
      )}
      <div
        style={{
          display: "flex",
          gap: 12,
          marginBottom: 18,
          flexWrap: "wrap",
        }}
      >
        <StatCard
          label="점검 학생"
          value={results.length}
          onClick={() => pickStatus("all")}
          active={statusFilter === "all"}
          hint="눌러서 전체 목록 보기"
        />
        <StatCard
          label="모든 요건 충족"
          value={passCount}
          tone={OK}
          onClick={() => pickStatus("pass")}
          active={statusFilter === "pass"}
        />
        <StatCard
          label="확인 필요"
          value={failCount}
          tone={WARN}
          onClick={() => pickStatus("fail")}
          active={statusFilter === "fail"}
          hint="눌러서 학생별 문제 보기"
        />
        <StatCard
          label="수강신청 기록이 빠진 학생"
          value={missingTotal}
          tone={missingTotal > 0 ? WARN : MUTED}
          onClick={() => pickStatus("missing")}
          active={statusFilter === "missing"}
        />
        <StatCard
          label="변경 후 시스템 미반영"
          value={unappliedTotal}
          tone={unappliedTotal > 0 ? WARN : MUTED}
          onClick={() => pickStatus("unapplied")}
          active={statusFilter === "unapplied"}
          hint="눌러서 반영할 학생 보기"
        />
      </div>
      {(statusFilter === "fail" || statusFilter === "unapplied") && (
        <ProblemListPanel
          mode={statusFilter}
          rows={filtered}
          selectedClass={selectedClass}
          onClearClass={() => setSelectedClass(null)}
          onClose={() => setStatusFilter("all")}
          onOpenDetail={setSelected}
          onOpenChange={onOpenChange}
          printOne={printOne}
          printFailOnly={() => printClass(selectedClass, true)}
          setChangeApplied={setChangeApplied}
          onPrintList={onPrintList}
          panelRef={panelRef}
        />
      )}
      {classSummary.length > 0 && (
        <div
          style={{
            ...CARD,
            marginBottom: 16,
            overflow: "hidden",
          }}
        >
          <div
            style={{
              padding: "14px 16px 8px",
            }}
          >
            <div
              style={{
                fontSize: 15.5,
                fontWeight: 800,
                color: INK,
              }}
            >
              학급별로 보기 (눌러서 필터)
            </div>
            <div
              style={{
                fontSize: 12,
                color: MUTED,
                marginTop: 3,
                lineHeight: 1.55,
              }}
            >
              ① 반별 인원이 실제 재적 인원과 같은지 ② ‘신청 누락’ 학생이 없는지 확인한 뒤 ③ 확인서를 반별로 인쇄하세요.
              ‘반 정보 없음’은 최근 학년 명단에 없는 학생(전출·자퇴 가능성, 또는 직접 추가 시 학번 미입력)입니다.
              {surveyMissingByClass && (
                <>
                  {" "}
                  ‘기초조사 안 냄’은 ② 반별 인원 중 아직 응답이 없는 번호입니다.
                  {surveyMissingOnlyClasses.length > 0 && (
                    <b
                      style={{
                        color: WARN,
                      }}
                    >
                      {" \uC751\uB2F5\uC774 \uD55C \uBA85\uB3C4 \uC5C6\uB294 \uBC18: "}
                      {surveyMissingOnlyClasses.join(", ")}
                    </b>
                  )}
                </>
              )}
            </div>
          </div>
          <div
            style={{
              overflowX: "auto",
            }}
          >
            <table
              style={{
                borderCollapse: "collapse",
                width: "100%",
                fontSize: 12.5,
                minWidth: 820,
              }}
            >
              <thead>
                <tr
                  style={{
                    background: PAPER,
                  }}
                >
                  <Th>학급</Th>
                  <Th>담당교사</Th>
                  <Th>인원</Th>
                  <Th>충족</Th>
                  <Th>확인필요</Th>
                  <Th>신청 누락</Th>
                  {surveyMissingByClass && <Th>기초조사 안 냄</Th>}
                  <Th>전입생</Th>
                  <Th>상담·변경</Th>
                  <Th>확인서 인쇄</Th>
                </tr>
              </thead>
              <tbody>
                <tr
                  style={{
                    borderTop: `1px solid ${LINE}`,
                  }}
                >
                  <td
                    onClick={() => setSelectedClass(null)}
                    style={{
                      padding: "9px 12px",
                      cursor: "pointer",
                      fontWeight: 700,
                      color: !selectedClass ? ACCENT : INK,
                      background: !selectedClass ? ACCENT_BG : "transparent",
                    }}
                  >
                    전체
                  </td>
                  <Td
                    style={{
                      color: MUTED,
                    }}
                  >
                    —
                  </Td>
                  <Td>{results.length}명</Td>
                  <Td
                    style={{
                      color: OK,
                    }}
                  >
                    {passCount}명
                  </Td>
                  <Td
                    style={{
                      color: WARN,
                    }}
                  >
                    {failCount}명
                  </Td>
                  <Td
                    style={{
                      color: missingTotal ? WARN : MUTED,
                    }}
                  >
                    {missingTotal}명
                  </Td>
                  {surveyMissingByClass && (
                    <Td
                      style={{
                        color: surveyNotSubmitted ? WARN : MUTED,
                        fontWeight: surveyNotSubmitted ? 700 : 400,
                      }}
                    >
                      {surveyNotSubmitted}명
                    </Td>
                  )}
                  <Td>{transferTotal}명</Td>
                  <Td>
                    {changedTotal}명
                    {unappliedTotal > 0 && (
                      <span
                        style={{
                          color: WARN,
                          fontWeight: 700,
                        }}
                      >
                        {" (\uBBF8\uBC18\uC601 "}
                        {unappliedTotal})
                      </span>
                    )}
                  </Td>
                  <Td>
                    <span
                      style={{
                        display: "flex",
                        gap: 4,
                      }}
                    >
                      <button onClick={() => printClass(null)} style={small("soft")}>
                        🖨 전체
                      </button>
                      <button onClick={() => printClass(null, true)} style={small("ghost")}>
                        확인필요만
                      </button>
                    </span>
                  </Td>
                </tr>
                {classSummary.map((c) => (
                  <tr
                    key={c.label}
                    style={{
                      borderTop: `1px solid ${LINE}`,
                      background: selectedClass === c.label ? ACCENT_BG : "transparent",
                    }}
                  >
                    <td
                      onClick={() => setSelectedClass(c.label)}
                      style={{
                        padding: "9px 12px",
                        cursor: "pointer",
                        fontWeight: 700,
                        color: selectedClass === c.label ? ACCENT : c.unclassified ? WARN : INK,
                        whiteSpace: "nowrap",
                      }}
                    >
                      {c.unclassified ? "반 정보 없음" : c.label}
                    </td>
                    <Td>
                      {c.unclassified ? (
                        <span
                          style={{
                            fontSize: 11.5,
                            color: MUTED,
                          }}
                        >
                          전출·명단 확인
                        </span>
                      ) : (
                        <input
                          value={classTeachers[c.label] || ""}
                          onChange={(e) =>
                            setClassTeachers((prev) => ({
                              ...prev,
                              [c.label]: e.target.value,
                            }))
                          }
                          placeholder="담임/진로담당"
                          style={{
                            width: 100,
                            fontSize: 12,
                            padding: "4px 6px",
                            borderRadius: 5,
                            border: `1px solid ${LINE}`,
                          }}
                        />
                      )}
                    </Td>
                    <Td>{c.count}명</Td>
                    <Td
                      style={{
                        color: OK,
                      }}
                    >
                      {c.passCount}명
                    </Td>
                    <Td
                      style={{
                        color: c.failCount > 0 ? WARN : MUTED,
                        fontWeight: c.failCount > 0 ? 700 : 400,
                      }}
                    >
                      {c.failCount}명
                    </Td>
                    <Td
                      style={{
                        color: c.missingCount > 0 ? WARN : MUTED,
                        fontWeight: c.missingCount > 0 ? 700 : 400,
                      }}
                    >
                      {c.missingCount}명
                    </Td>
                    {surveyMissingByClass && (
                      <Td
                        style={{
                          color: (surveyMissingByClass.get(c.label) || []).length ? WARN : MUTED,
                          whiteSpace: "normal",
                          maxWidth: 220,
                        }}
                      >
                        {c.unclassified
                          ? "—"
                          : (surveyMissingByClass.get(c.label) || []).length
                            ? (surveyMissingByClass.get(c.label) || []).map((n) => `${n}번`).join(", ")
                            : "없음"}
                      </Td>
                    )}
                    <Td
                      style={{
                        color: c.transferCount ? INK : MUTED,
                      }}
                    >
                      {c.transferCount}명
                    </Td>
                    <Td
                      style={{
                        color: c.changedCount ? INK : MUTED,
                      }}
                    >
                      {c.changedCount}명
                      {c.unappliedCount > 0 && (
                        <span
                          style={{
                            color: WARN,
                            fontWeight: 700,
                          }}
                        >
                          {" (\uBBF8\uBC18\uC601 "}
                          {c.unappliedCount})
                        </span>
                      )}
                    </Td>
                    <Td>
                      <span
                        style={{
                          display: "flex",
                          gap: 4,
                        }}
                      >
                        <button onClick={() => printClass(c.label)} style={small("soft")}>
                          🖨 학급
                        </button>
                        <button onClick={() => printClass(c.label, true)} style={small("ghost")}>
                          확인필요만
                        </button>
                      </span>
                    </Td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
      <div
        style={{
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          marginBottom: 10,
          gap: 12,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            gap: 8,
            flexWrap: "wrap",
          }}
        >
          <div
            style={{
              position: "relative",
              width: 240,
            }}
          >
            <Search
              size={14}
              style={{
                position: "absolute",
                left: 10,
                top: 10,
                color: MUTED,
              }}
            />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="학번, 이름 검색"
              style={{
                width: "100%",
                padding: "8px 10px 8px 30px",
                borderRadius: 7,
                border: `1px solid ${LINE}`,
                fontSize: 13,
                boxSizing: "border-box",
              }}
            />
          </div>
          {statusFilters.map((f) => (
            <button key={f.id} onClick={() => setStatusFilter(f.id)} style={chipStyle(statusFilter === f.id)}>
              {f.label} {f.count}
            </button>
          ))}
          {selectedClass && (
            <button onClick={() => setSelectedClass(null)} style={chipStyle(true)}>
              {selectedClass === UNCLASSIFIED_LABEL ? "반 정보 없음" : selectedClass}
              {" \u2715"}
            </button>
          )}
        </div>
        <button
          onClick={handleExport}
          style={buttonStyle("ghost", {
            color: ACCENT,
            borderColor: ACCENT,
            padding: "8px 14px",
          })}
        >
          <Download size={14} />
          {" \uC810\uAC80\uACB0\uACFC CSV \uB2E4\uC6B4\uB85C\uB4DC"}
        </button>
      </div>
      <div
        ref={tableRef}
        style={{
          ...CARD,
          overflow: "auto",
          scrollMarginTop: 120,
        }}
      >
        <table
          style={{
            borderCollapse: "collapse",
            width: "100%",
            fontSize: 12.5,
            minWidth: 980,
          }}
        >
          <thead>
            <tr
              style={{
                background: PAPER,
              }}
            >
              <Th>학번</Th>
              <Th>이름</Th>
              {groups.map((g) => (
                <th
                  key={g}
                  style={{
                    textAlign: "left",
                    padding: "10px 12px",
                    fontSize: 12,
                    fontWeight: 700,
                    color: priorityGroups[g] ? "#8A6D00" : MUTED,
                    whiteSpace: "nowrap",
                    position: "sticky",
                    top: 0,
                    background: priorityGroups[g] ? "#FDF3C7" : PAPER,
                  }}
                >
                  {priorityGroups[g] ? "⭐ " : ""}
                  {g}
                </th>
              ))}
              <Th>총학점</Th>
              <Th>판정</Th>
              <Th>확인 사유</Th>
              <Th>보관</Th>
            </tr>
          </thead>
          <tbody>
            {filtered.length === 0 && (
              <tr>
                <td
                  colSpan={groups.length + 6}
                  style={{
                    padding: 20,
                    textAlign: "center",
                    color: MUTED,
                  }}
                >
                  조건에 맞는 학생이 없습니다.
                </td>
              </tr>
            )}
            {filtered.map((r) => (
              <tr
                key={r.key}
                onClick={() => setSelected(r.key)}
                style={{
                  borderTop: `1px solid ${LINE}`,
                  cursor: "pointer",
                  background: selected === r.key ? ACCENT_BG : "transparent",
                }}
              >
                <Td
                  style={{
                    color: MUTED,
                  }}
                >
                  {r.studentId}
                </Td>
                <Td
                  style={{
                    fontWeight: 600,
                  }}
                >
                  {r.name}
                  <StudentTags r={r} />
                </Td>
                {groups.map((g) => {
                  var _a, _b;
                  const cur = r.groupCredit[g] || 0;
                  const short = cur < ((_a = reqOverride[g]) !== null && _a !== void 0 ? _a : 0);
                  return (
                    <Td
                      key={g}
                      style={{
                        color: short ? WARN : INK,
                        fontWeight: short ? 700 : 400,
                        background: priorityGroups[g] ? "#FFFCF0" : "transparent",
                      }}
                    >
                      {cur}/{(_b = reqOverride[g]) !== null && _b !== void 0 ? _b : 0}
                    </Td>
                  );
                })}
                <Td
                  style={{
                    fontWeight: 700,
                    color: r.total < minCredit ? WARN : INK,
                  }}
                >
                  {r.total}
                </Td>
                <Td>
                  <PassBadge pass={r.pass} />
                </Td>
                <Td
                  style={{
                    whiteSpace: "normal",
                    minWidth: 150,
                    maxWidth: 260,
                  }}
                >
                  <ReasonList reasons={r.reasons.slice(0, 3)} compact={true} />
                  {r.reasons.length > 3 && (
                    <span
                      style={{
                        fontSize: 11,
                        color: MUTED,
                      }}
                    >
                      {"\uC678 "}
                      {r.reasons.length - 3}
                    </span>
                  )}
                </Td>
                <Td>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      setArchiveModalKey(r.key);
                    }}
                    style={small("ghost")}
                  >
                    보관
                  </button>
                </Td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div
        style={{
          fontSize: 12,
          color: MUTED,
          marginTop: 10,
          lineHeight: 1.55,
        }}
      >
        행을 누르면 학생별 상세 내역이 열립니다. 상담 후 바뀐 선택과목은 ④, 전입생의 전적교 과목은 ⑤, 공동교육과정 등은
        ⑥에서 입력하면 이 표와 확인서에 바로 반영됩니다. ‘확인 사유’의 빨간 항목이 판정을 ‘확인필요’로 만든 원인이고,
        회색 항목은 참고용입니다.
      </div>
      {archivedResults.length > 0 && (
        <div
          style={{
            ...CARD,
            marginTop: 16,
            overflow: "hidden",
          }}
        >
          <button
            onClick={() => setShowArchived((v) => !v)}
            style={{
              width: "100%",
              display: "flex",
              alignItems: "center",
              justifyContent: "space-between",
              padding: "12px 16px",
              background: "none",
              border: "none",
              cursor: "pointer",
              fontSize: 13,
              fontWeight: 700,
              color: INK,
            }}
          >
            <span>보관된 학생 ({archivedResults.length}명)</span>
            {showArchived ? <ChevronUp size={15} /> : <ChevronDown size={15} />}
          </button>
          {showArchived && (
            <table
              style={{
                borderCollapse: "collapse",
                width: "100%",
                fontSize: 12.5,
              }}
            >
              <thead>
                <tr
                  style={{
                    background: PAPER,
                  }}
                >
                  <Th>학번</Th>
                  <Th>이름</Th>
                  <Th>사유</Th>
                  <Th>복원</Th>
                </tr>
              </thead>
              <tbody>
                {archivedResults.map((r) => (
                  <tr
                    key={r.key}
                    style={{
                      borderTop: `1px solid ${LINE}`,
                    }}
                  >
                    <Td
                      style={{
                        color: MUTED,
                      }}
                    >
                      {r.studentId}
                    </Td>
                    <Td
                      style={{
                        fontWeight: 600,
                      }}
                    >
                      {r.name}
                    </Td>
                    <Td
                      style={{
                        color: MUTED,
                      }}
                    >
                      {r.archiveInfo.reason || "(사유 없음)"}
                    </Td>
                    <Td>
                      <button onClick={() => restoreStudent(r.key)} style={small("soft")}>
                        복원
                      </button>
                    </Td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
      {archiveModalKey && (
        <ModalShell
          onClose={() => {
            setArchiveModalKey(null);
            setArchiveReason("");
          }}
        >
          <div
            style={{
              fontSize: 15,
              fontWeight: 700,
              marginBottom: 10,
            }}
          >
            학생 보관
          </div>
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              marginBottom: 12,
            }}
          >
            {
              '\uC804\uD559\u00B7\uC790\uD1F4 \uB4F1\uC758 \uC0AC\uC720\uB85C \uC810\uAC80 \uBAA9\uB85D\uC5D0\uC11C \uB530\uB85C \uBE7C\uB450\uB824\uBA74 \uC0AC\uC720\uB97C \uC801\uC5B4\uC8FC\uC138\uC694. \uBCF4\uAD00\uB41C \uD559\uC0DD\uC740 \uC544\uB798 "\uBCF4\uAD00\uB41C \uD559\uC0DD" \uBAA9\uB85D\uC5D0\uC11C \uC5B8\uC81C\uB4E0 \uBCF5\uC6D0\uD560 \uC218 \uC788\uC2B5\uB2C8\uB2E4.'
            }
          </div>
          <textarea
            value={archiveReason}
            onChange={(e) => setArchiveReason(e.target.value)}
            placeholder="예: 2026.3 전학"
            rows={3}
            style={{
              width: "100%",
              boxSizing: "border-box",
              padding: "9px 11px",
              borderRadius: 7,
              border: `1px solid ${LINE}`,
              fontSize: 13,
              resize: "vertical",
              marginBottom: 14,
            }}
          />
          <div
            style={{
              display: "flex",
              justifyContent: "flex-end",
              gap: 8,
            }}
          >
            <button
              onClick={() => {
                setArchiveModalKey(null);
                setArchiveReason("");
              }}
              style={{
                fontSize: 13,
                fontWeight: 600,
                color: MUTED,
                background: "none",
                border: "none",
                padding: "8px 12px",
                cursor: "pointer",
              }}
            >
              취소
            </button>
            <button
              onClick={confirmArchive}
              style={buttonStyle("primary", {
                fontSize: 13,
                padding: "8px 16px",
              })}
            >
              보관
            </button>
          </div>
        </ModalShell>
      )}
      {selectedResult && (
        <div
          onClick={() => setSelected(null)}
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(28,35,51,0.35)",
            display: "flex",
            justifyContent: "flex-end",
            zIndex: 50,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{
              width: 470,
              maxWidth: "92vw",
              background: "#fff",
              height: "100%",
              overflowY: "auto",
              padding: 24,
              fontFamily: FONT,
              boxShadow: "-8px 0 24px rgba(0,0,0,0.12)",
              boxSizing: "border-box",
            }}
          >
            <StudentDetail
              r={selectedResult}
              groups={groups}
              reqOverride={reqOverride}
              priorityGroups={priorityGroups}
              curriculumFile={curriculumFile}
              onClose={() => setSelected(null)}
              printOne={printOne}
              onOpenChange={onOpenChange}
              onOpenTransfer={onOpenTransfer}
              onOpenExtra={onOpenExtra}
              setChangeApplied={setChangeApplied}
            />
          </div>
        </div>
      )}
    </div>
  );
}
function StudentDetail({
  r,
  groups,
  reqOverride,
  priorityGroups,
  curriculumFile,
  onClose,
  printOne,
  onOpenChange,
  onOpenTransfer,
  onOpenExtra,
  setChangeApplied,
}) {
  const small = (kind) =>
    buttonStyle(kind, {
      fontSize: 11.5,
      padding: "4px 9px",
    });
  const warnBox = {
    marginTop: 14,
    background: WARN_BG,
    borderRadius: 8,
    padding: "10px 12px",
  };
  return (
    <div>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          gap: 8,
        }}
      >
        <div>
          <div
            style={{
              fontSize: 18,
              fontWeight: 700,
            }}
          >
            {r.name}
            <StudentTags r={r} />
          </div>
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              marginTop: 2,
            }}
          >
            {r.studentId}
            {" \u00B7 "}
            {r.classLabel}
          </div>
        </div>
        <div
          style={{
            display: "flex",
            gap: 6,
          }}
        >
          <button
            onClick={() => printOne(r.key)}
            style={buttonStyle("ghost", {
              fontSize: 12,
              padding: "6px 10px",
            })}
          >
            🖨 확인서
          </button>
          <button
            onClick={onClose}
            style={{
              border: "none",
              background: "none",
              cursor: "pointer",
              color: MUTED,
            }}
          >
            <X size={18} />
          </button>
        </div>
      </div>
      <div
        style={{
          marginTop: 14,
          display: "inline-block",
          fontSize: 12.5,
          fontWeight: 700,
          padding: "4px 11px",
          borderRadius: 999,
          background: r.pass ? OK_BG : WARN_BG,
          color: r.pass ? OK : WARN,
        }}
      >
        {r.pass ? "졸업요건 충족" : "확인 필요"}
        {" \u00B7 \uCD1D "}
        {r.total}학점
      </div>
      <div
        style={{
          display: "flex",
          gap: 6,
          flexWrap: "wrap",
          marginTop: 10,
        }}
      >
        <button onClick={() => onOpenChange(r.key)} style={small("soft")}>
          ④ 선택과목 변경
        </button>
        <button onClick={() => onOpenTransfer(r.key)} style={small("soft")}>
          {"\u2464 "}
          {r.isTransfer ? "전입생 정보 수정" : "전입생으로 등록"}
        </button>
        <button onClick={onOpenExtra} style={small("soft")}>
          ⑥ 공동교육과정 등
        </button>
      </div>
      {r.reasons.length > 0 && (
        <div
          style={{
            marginTop: 12,
          }}
        >
          <ReasonList reasons={r.reasons} />
        </div>
      )}
      {r.missingSemesters.length > 0 && (
        <div style={warnBox}>
          <div
            style={{
              fontSize: 12.5,
              fontWeight: 700,
              color: WARN,
              marginBottom: 4,
            }}
          >
            {"\u26A0 \uC218\uAC15\uC2E0\uCCAD \uAE30\uB85D\uC774 \uC5C6\uB294 \uD559\uAE30: "}
            {r.missingSemesters.map(formatSemester).join(", ")}
          </div>
          <div
            style={{
              fontSize: 11.5,
              color: WARN,
              lineHeight: 1.55,
            }}
          >
            그 학기의 수강신청 파일은 올라왔는데 이 학생만 없습니다. 파일에서 빠졌는지, 동명이인 병합이 필요한지,
            전입생인지 확인하세요. 실제 선택 과목은 ④에서 입력할 수 있습니다.
          </div>
        </div>
      )}
      {r.duplicateSelections.length > 0 && (
        <div style={warnBox}>
          <div
            style={{
              fontSize: 12.5,
              fontWeight: 700,
              color: WARN,
              marginBottom: 4,
            }}
          >
            ⚠ 같은 학년 1·2학기에 걸쳐 선택된 과목
          </div>
          <div
            style={{
              fontSize: 12,
              color: WARN,
              lineHeight: 1.6,
            }}
          >
            {r.duplicateSelections.map((d, i) => (
              <div key={i}>
                {d.name}
                {" \u2014 "}
                {d.semesters.map(formatSemester).join(", ")}에 모두 선택됨
              </div>
            ))}
          </div>
          <div
            style={{
              fontSize: 11,
              color: WARN,
              marginTop: 4,
              opacity: 0.85,
            }}
          >
            같은 과목을 두 학기에 걸쳐 이수할 수 없습니다. ④에서 한쪽을 취소하세요.
          </div>
        </div>
      )}
      <div
        style={{
          marginTop: 20,
        }}
      >
        <div
          style={{
            fontSize: 13,
            fontWeight: 700,
            marginBottom: 8,
          }}
        >
          교과영역별 이수학점
        </div>
        <GroupCreditGrid result={r} groups={groups} reqOverride={reqOverride} priorityGroups={priorityGroups} />
      </div>
      {r.shortGroups.length > 0 && curriculumFile && (
        <div
          style={{
            marginTop: 20,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 8,
              color: WARN,
            }}
          >
            부족한 영역에서 들을 수 있는 과목
          </div>
          {r.shortGroups.map((g) => {
            const takenKeys = new Set(r.taken.map((t) => normalizeName(t.name)));
            const seen = new Set();
            const options = curriculumFile.courses
              .filter(
                (c) =>
                  toCurriculumGroup(c.group) === g &&
                  c.gubun === "학생선택" &&
                  !takenKeys.has(c.key) &&
                  isAtThisSchool(c.semester, r.firstSem),
              )
              .filter((c) => (seen.has(c.key) ? false : (seen.add(c.key), true)));
            return (
              <div
                key={g}
                style={{
                  marginBottom: 12,
                }}
              >
                <div
                  style={{
                    fontSize: 12.5,
                    fontWeight: 700,
                    marginBottom: 4,
                  }}
                >
                  {g}
                </div>
                {options.length === 0 ? (
                  <div
                    style={{
                      fontSize: 12,
                      color: MUTED,
                    }}
                  >
                    편제표에 등록된 추가 선택과목이 없습니다.
                  </div>
                ) : (
                  options.map((c, i) => (
                    <div
                      key={i}
                      style={{
                        display: "flex",
                        justifyContent: "space-between",
                        fontSize: 12,
                        padding: "3px 2px",
                      }}
                    >
                      <span>{c.name}</span>
                      <span
                        style={{
                          color: MUTED,
                        }}
                      >
                        {formatSemester(c.semester)}
                        {" \u00B7 "}
                        {c.credit}학점
                      </span>
                    </div>
                  ))
                )}
              </div>
            );
          })}
        </div>
      )}
      {r.pools.length > 0 && (
        <div
          style={{
            marginTop: 20,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 8,
            }}
          >
            선택과목 이수 현황 (학기 · 택N별)
          </div>
          {Array.from(new Set(r.pools.map((p) => p.sem))).map((sem) => (
            <div
              key={sem}
              style={{
                marginBottom: 12,
              }}
            >
              <div
                style={{
                  fontSize: 12.5,
                  fontWeight: 700,
                  color: INK,
                  marginBottom: 6,
                }}
              >
                {formatSemester(sem)}
              </div>
              {r.pools
                .filter((p) => p.sem === sem)
                .map((p, pi) => (
                  <div
                    key={pi}
                    style={{
                      border: `1px solid ${p.deficient || p.over ? WARN : LINE}`,
                      background: p.deficient || p.over ? WARN_BG : PAPER,
                      borderRadius: 8,
                      padding: "8px 10px",
                      marginBottom: 6,
                    }}
                  >
                    <div
                      style={{
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "space-between",
                        marginBottom: p.takenCourses.length ? 6 : 0,
                      }}
                    >
                      <span
                        style={{
                          fontSize: 12,
                          fontWeight: 700,
                          color: p.deficient || p.over ? WARN : INK,
                        }}
                      >
                        {p.label}
                      </span>
                      <span
                        style={{
                          fontSize: 11,
                          fontWeight: 700,
                          color: p.deficient || p.over ? WARN : OK,
                        }}
                      >
                        {p.takenCount}
                        {p.requiredN != null ? `/${p.requiredN}` : ""}
                        {" \uC120\uD0DD"}
                      </span>
                    </div>
                    {p.takenCourses.map((c, ci) => (
                      <div
                        key={ci}
                        style={{
                          display: "flex",
                          justifyContent: "space-between",
                          fontSize: 12,
                          padding: "2px 0",
                        }}
                      >
                        <span>
                          <span
                            style={{
                              color: MUTED,
                            }}
                          >
                            [{c.group}]
                          </span>{" "}
                          {c.name}
                        </span>
                        <span>{c.credit}</span>
                      </div>
                    ))}
                    {p.takenCourses.length === 0 && p.deficient && (
                      <div
                        style={{
                          fontSize: 12,
                          color: WARN,
                        }}
                      >
                        선택한 과목 없음
                      </div>
                    )}
                  </div>
                ))}
            </div>
          ))}
        </div>
      )}
      {r.semesterMismatch.length > 0 && (
        <div
          style={{
            ...warnBox,
            marginTop: 20,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 6,
              color: WARN,
            }}
          >
            편제표의 개설 학기와 다른 학기에 신청된 과목
          </div>
          {r.semesterMismatch.map((d, i) => (
            <div
              key={i}
              style={{
                fontSize: 12.5,
                color: WARN,
              }}
            >
              {d.name}
              {" \u2014 "}
              {formatSemester(d.semester)}
              {"\uC5D0 \uC2E0\uCCAD \u00B7 \uD3B8\uC81C\uD45C\uC5D0\uB294 "}
              {d.offered.map(formatSemester).join(", ")}
              {" \uAC1C\uC124"}
            </div>
          ))}
          <div
            style={{
              fontSize: 11,
              color: WARN,
              marginTop: 4,
              opacity: 0.85,
            }}
          >
            학점은 신청한 학기 기준으로 반영했고, 택N 선택 현황에서는 제외했습니다.
          </div>
        </div>
      )}
      {r.notInCurriculum.length > 0 && (
        <div
          style={{
            marginTop: 16,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 6,
              color: ACCENT,
            }}
          >
            학교 편제표엔 없지만 과목목록으로 인식된 과목 (학점 반영)
          </div>
          {r.notInCurriculum.map((n, i) => (
            <div
              key={i}
              style={{
                fontSize: 12.5,
                color: ACCENT,
              }}
            >
              {n}
            </div>
          ))}
        </div>
      )}
      {r.unmatched.length > 0 && (
        <div
          style={{
            marginTop: 16,
          }}
        >
          <div
            style={{
              fontSize: 13,
              fontWeight: 700,
              marginBottom: 6,
              color: WARN,
            }}
          >
            편제표에 없는 과목 (학점 미반영)
          </div>
          {r.unmatched.map((n, i) => (
            <div
              key={i}
              style={{
                fontSize: 12.5,
                color: WARN,
              }}
            >
              {n}
            </div>
          ))}
        </div>
      )}
      <DetailSection
        title="④ 상담 후 선택과목 변경"
        action={
          <button onClick={() => onOpenChange(r.key)} style={small("soft")}>
            {r.hasChanges ? "수정" : "변경하기"}
          </button>
        }
      >
        {r.change.added.length === 0 && r.change.removed.length === 0 && !r.change.memo ? (
          <div
            style={{
              fontSize: 12,
              color: MUTED,
            }}
          >
            변경 내역이 없습니다.
          </div>
        ) : (
          <div
            style={{
              fontSize: 12.5,
              lineHeight: 1.7,
            }}
          >
            {r.change.added.map((e, i) => (
              <div
                key={`a${i}`}
                style={{
                  color: OK,
                }}
              >
                {"+ \uCD94\uAC00 \u00B7 "}
                {e.name}
                {" ("}
                {formatSemester(e.semester)})
              </div>
            ))}
            {r.change.removed.map((e, i) => (
              <div
                key={`r${i}`}
                style={{
                  color: WARN,
                }}
              >
                {"\u2212 \uCDE8\uC18C \u00B7 "}
                {e.name}
                {" ("}
                {formatSemester(e.semester)})
              </div>
            ))}
            {r.change.memo && (
              <div
                style={{
                  color: MUTED,
                  whiteSpace: "pre-wrap",
                  marginTop: 4,
                }}
              >
                {"\uC0C1\uB2F4 \uBA54\uBAA8: "}
                {r.change.memo}
              </div>
            )}
            {r.change.applyState !== "none" && setChangeApplied && (
              <div
                style={{
                  marginTop: 6,
                }}
              >
                <ChangeApplyBox r={r} setChangeApplied={setChangeApplied} compact={true} />
              </div>
            )}
          </div>
        )}
      </DetailSection>
      <DetailSection
        title="⑤ 전입생 정보"
        action={
          <button onClick={() => onOpenTransfer(r.key)} style={small("soft")}>
            {r.isTransfer ? "수정" : "전입생으로 등록"}
          </button>
        }
      >
        {!r.isTransfer ? (
          <div
            style={{
              fontSize: 12,
              color: MUTED,
            }}
          >
            전입생이 아닙니다. 이전 학기를 다른 학교에서 이수했다면 전입생으로 등록하세요.
          </div>
        ) : (
          <div
            style={{
              fontSize: 12.5,
              lineHeight: 1.65,
            }}
          >
            <div>
              {"\uBCF8\uAD50 \uCCAB \uD559\uAE30 "}
              <b>{formatSemester(r.firstSem)}</b>
              {r.prevSchool && ` · 전적교 ${r.prevSchool}`}
            </div>
            {r.priorCourses.length === 0 ? (
              <div
                style={{
                  color: WARN,
                }}
              >
                전입 이전 학기 이수 과목이 아직 입력되지 않았습니다.
              </div>
            ) : (
              r.priorCourses.map((p, i) => (
                <div
                  key={i}
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                  }}
                >
                  <span>
                    <span
                      style={{
                        color: MUTED,
                      }}
                    >
                      {formatSemester(p.semester)}
                      {" ["}
                      {p.group || "교과군 없음"}]
                    </span>{" "}
                    {p.name}
                  </span>
                  <span>{p.credit}</span>
                </div>
              ))
            )}
            {r.priorIgnored.length > 0 && (
              <div
                style={{
                  color: WARN,
                }}
              >
                {"\uBCF8\uAD50 \uC7AC\uD559 \uD559\uAE30\uB85C \uC785\uB825\uB41C \uC804\uC801\uAD50 \uACFC\uBAA9 "}
                {r.priorIgnored.length}개는 계산에서 뺐습니다.
              </div>
            )}
            {r.excludedBeforeTransfer.length > 0 && (
              <div
                style={{
                  color: MUTED,
                }}
              >
                {"\uC804\uC785 \uC774\uC804 \uD559\uAE30\uC758 \uC218\uAC15\uC2E0\uCCAD \uAE30\uB85D "}
                {r.excludedBeforeTransfer.length}과목은 계산에서 뺐습니다.
              </div>
            )}
          </div>
        )}
      </DetailSection>
      <DetailSection
        title="⑥ 공동교육과정 등 추가 이수"
        action={
          <button onClick={onOpenExtra} style={small("soft")}>
            ⑥으로 이동
          </button>
        }
      >
        {r.extraEntries.length === 0 ? (
          <div
            style={{
              fontSize: 12,
              color: MUTED,
            }}
          >
            추가 이수 과목이 없습니다.
          </div>
        ) : (
          r.extraEntries.map((x) => (
            <div
              key={x.id}
              style={{
                display: "flex",
                justifyContent: "space-between",
                fontSize: 12.5,
                padding: "2px 0",
              }}
            >
              <span>
                <span
                  style={{
                    color: MUTED,
                  }}
                >
                  {x.kind}
                  {" ["}
                  {x.group || "교과군 없음"}]
                </span>{" "}
                {x.name}
                {x.semester && (
                  <span
                    style={{
                      color: MUTED,
                    }}
                  >
                    {" \u00B7 "}
                    {formatSemester(x.semester)}
                  </span>
                )}
              </span>
              <span
                style={{
                  fontWeight: 700,
                }}
              >
                {x.credit}학점
              </span>
            </div>
          ))
        )}
      </DetailSection>
    </div>
  );
}
// ================= PAGE 4: 선택과목 변경 (상담 후 수정) =================
function CourseToggle({ state, name, group, credit, onClick }) {
  const styles = {
    on: {
      border: `1.5px solid ${ACCENT}`,
      background: ACCENT_BG,
      color: ACCENT,
    },
    added: {
      border: `1.5px solid ${OK}`,
      background: OK_BG,
      color: OK,
    },
    removed: {
      border: `1.5px dashed ${WARN}`,
      background: "#fff",
      color: WARN,
    },
    off: {
      border: `1px solid ${LINE}`,
      background: "#fff",
      color: INK,
    },
  };
  const mark = {
    on: "✓",
    added: "✓ 추가",
    removed: "취소됨",
    off: "+",
  }[state];
  return (
    <button
      onClick={onClick}
      title={
        state === "removed" ? "다시 누르면 취소를 되돌립니다" : state === "added" ? "다시 누르면 추가를 되돌립니다" : ""
      }
      style={{
        ...styles[state],
        borderRadius: 8,
        padding: "6px 9px",
        cursor: "pointer",
        textAlign: "left",
        fontFamily: FONT,
        minWidth: 0,
      }}
    >
      <div
        style={{
          fontSize: 12.5,
          fontWeight: 700,
          textDecoration: state === "removed" ? "line-through" : "none",
        }}
      >
        <span
          style={{
            marginRight: 5,
          }}
        >
          {mark}
        </span>
        {name}
      </div>
      <div
        style={{
          fontSize: 11,
          color: MUTED,
        }}
      >
        {group}
        {" \u00B7 "}
        {credit}학점
      </div>
    </button>
  );
}
function ChangeEditor({
  r,
  groups,
  reqOverride,
  priorityGroups,
  poolDefs,
  toggleCourseChange,
  setChangeMemo,
  resetCourseChange,
  onOpenCheck,
  onOpenTransfer,
  printOne,
  setChangeApplied,
}) {
  const currentKeys = new Set(r.entries.map((e) => entryKey(e.name, e.semester)));
  const registeredKeys = new Set(r.registered.map((e) => entryKey(e.name, e.semester)));
  const hereDefs = poolDefs.filter((d) => isAtThisSchool(d.sem, r.firstSem));
  const sems = Array.from(new Set(hereDefs.map((d) => d.sem))).sort(compareSemester);
  const poolKeySet = new Set(hereDefs.flatMap((d) => d.courses.map((c) => entryKey(c.name, d.sem))));
  const poolByKey = new Map(r.pools.map((p) => [`${p.sem}||${p.label}`, p]));
  const others = [
    ...r.entries
      .filter((e) => !poolKeySet.has(entryKey(e.name, e.semester)))
      .map((e) => ({
        ...e,
        on: true,
      })),
    ...r.change.removed
      .filter((e) => !poolKeySet.has(entryKey(e.name, e.semester)))
      .map((e) => ({
        ...e,
        on: false,
      })),
  ].sort((a, b) => compareSemester(a.semester, b.semester));
  const takenByKey = new Map(r.taken.map((t) => [entryKey(t.name, t.semester), t]));
  const changedCount = r.change.added.length + r.change.removed.length;
  return (
    <div
      style={{
        ...CARD,
        padding: 18,
        minWidth: 0,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          gap: 10,
          flexWrap: "wrap",
        }}
      >
        <div>
          <div
            style={{
              fontSize: 18,
              fontWeight: 800,
            }}
          >
            {r.name}
            <StudentTags r={r} />
          </div>
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              marginTop: 2,
            }}
          >
            {r.studentId}
            {" \u00B7 "}
            {r.classLabel}
          </div>
        </div>
        <div
          style={{
            display: "flex",
            gap: 6,
            alignItems: "center",
            flexWrap: "wrap",
          }}
        >
          <PassBadge pass={r.pass} size={12.5} />
          <span
            style={{
              fontSize: 13,
              fontWeight: 700,
            }}
          >
            {"\uCD1D "}
            {r.total}학점
          </span>
          <button
            onClick={() => onOpenCheck(r.key)}
            style={buttonStyle("ghost", {
              fontSize: 12,
            })}
          >
            ③ 상세 보기
          </button>
          <button
            onClick={() => printOne(r.key)}
            style={buttonStyle("ghost", {
              fontSize: 12,
            })}
          >
            🖨 확인서
          </button>
        </div>
      </div>
      {r.reasons.length > 0 ? (
        <div
          style={{
            marginTop: 10,
          }}
        >
          <ReasonList reasons={r.reasons} />
        </div>
      ) : (
        <div
          style={{
            marginTop: 10,
            fontSize: 12.5,
            color: OK,
            fontWeight: 700,
          }}
        >
          ✓ 확인이 필요한 항목이 없습니다.
        </div>
      )}
      <div
        style={{
          marginTop: 12,
        }}
      >
        <GroupCreditGrid result={r} groups={groups} reqOverride={reqOverride} priorityGroups={priorityGroups} />
      </div>
      {r.isTransfer && (
        <div
          style={{
            marginTop: 10,
            fontSize: 12,
            color: MUTED,
          }}
        >
          {"\uC804\uC785\uC0DD \u2014 \uBCF8\uAD50 "}
          {formatSemester(r.firstSem)}부터의 선택과목만 여기서 고칩니다. 전입 이전 과목은{" "}
          <button
            onClick={() => onOpenTransfer(r.key)}
            style={{
              border: "none",
              background: "none",
              color: ACCENT,
              fontWeight: 700,
              cursor: "pointer",
              padding: 0,
            }}
          >
            ⑤ 전입생
          </button>
          에서 입력하세요.
        </div>
      )}
      {sems.length === 0 && (
        <div
          style={{
            marginTop: 16,
            fontSize: 12.5,
            color: MUTED,
          }}
        >
          편제표에 이 학생이 고를 수 있는 선택그룹(택N)이 없습니다.
        </div>
      )}
      {sems.map((sem) => (
        <div
          key={sem}
          style={{
            marginTop: 18,
          }}
        >
          <div
            style={{
              fontSize: 14,
              fontWeight: 800,
              marginBottom: 6,
            }}
          >
            {formatSemester(sem)}
            {r.missingSemesters.includes(sem) && (
              <span
                style={{
                  marginLeft: 8,
                  fontSize: 11.5,
                  fontWeight: 700,
                  color: WARN,
                  background: WARN_BG,
                  borderRadius: 999,
                  padding: "2px 8px",
                }}
              >
                수강신청 기록 없음
              </span>
            )}
          </div>
          {hereDefs
            .filter((d) => d.sem === sem)
            .map((d) => {
              const p = poolByKey.get(`${d.sem}||${d.label}`);
              const n = d.requiredN;
              const cnt = p ? p.takenCount : 0;
              const tone = n == null ? MUTED : cnt === n ? OK : WARN;
              return (
                <div
                  key={`${d.sem}||${d.label}`}
                  style={{
                    border: `1px solid ${LINE}`,
                    borderRadius: 9,
                    padding: 10,
                    marginBottom: 8,
                    background: PAPER,
                  }}
                >
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      marginBottom: 8,
                      gap: 8,
                      flexWrap: "wrap",
                    }}
                  >
                    <span
                      style={{
                        fontSize: 12.5,
                        fontWeight: 700,
                      }}
                    >
                      {d.label || "선택그룹"}
                    </span>
                    <span
                      style={{
                        fontSize: 12,
                        fontWeight: 700,
                        color: tone,
                      }}
                    >
                      {cnt}
                      {n != null ? ` / ${n}` : ""}
                      {" \uC120\uD0DD"}
                      {n != null && cnt < n && ` · ${n - cnt}개 더 골라야 함`}
                      {n != null && cnt > n && ` · ${cnt - n}개 초과`}
                    </span>
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(auto-fill, minmax(170px, 1fr))",
                      gap: 6,
                    }}
                  >
                    {d.courses.map((c) => {
                      const k = entryKey(c.name, d.sem);
                      const on = currentKeys.has(k);
                      const orig = registeredKeys.has(k);
                      const state = on && !orig ? "added" : !on && orig ? "removed" : on ? "on" : "off";
                      return (
                        <CourseToggle
                          key={k}
                          state={state}
                          name={c.name}
                          group={toCurriculumGroup(c.group)}
                          credit={c.credit}
                          onClick={() => toggleCourseChange(r, c.name, d.sem)}
                        />
                      );
                    })}
                  </div>
                </div>
              );
            })}
        </div>
      ))}
      {others.length > 0 && (
        <div
          style={{
            marginTop: 18,
          }}
        >
          <div
            style={{
              fontSize: 13.5,
              fontWeight: 700,
              marginBottom: 4,
            }}
          >
            선택그룹 밖에서 신청된 과목
          </div>
          <div
            style={{
              fontSize: 12,
              color: MUTED,
              marginBottom: 8,
            }}
          >
            편제표의 선택그룹에 없거나 개설 학기가 다른 과목입니다. 실제로 듣지 않는 과목이면 눌러서 취소하세요.
          </div>
          <div
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fill, minmax(190px, 1fr))",
              gap: 6,
            }}
          >
            {others.map((e) => {
              const k = entryKey(e.name, e.semester);
              const t = takenByKey.get(k);
              const state = e.on ? (e.origin === "added" ? "added" : "on") : "removed";
              return (
                <CourseToggle
                  key={k}
                  state={state}
                  name={`${e.name} (${formatSemester(e.semester)})`}
                  group={t ? t.group : "편제표에 없음"}
                  credit={t ? t.credit : 0}
                  onClick={() => toggleCourseChange(r, e.name, e.semester)}
                />
              );
            })}
          </div>
        </div>
      )}
      <div
        style={{
          marginTop: 20,
          borderTop: `1px solid ${LINE}`,
          paddingTop: 14,
        }}
      >
        <div
          style={{
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            gap: 8,
            flexWrap: "wrap",
            marginBottom: 6,
          }}
        >
          <div
            style={{
              fontSize: 13.5,
              fontWeight: 700,
            }}
          >
            {"\uBCC0\uACBD \uB0B4\uC5ED "}
            {changedCount > 0 && `(${changedCount}건)`}
          </div>
          {(changedCount > 0 || r.change.memo) && (
            <button
              onClick={() =>
                window.confirm(`${r.name} 학생의 변경 내역과 상담 메모를 모두 지우고 원래 수강신청대로 되돌릴까요?`) &&
                resetCourseChange(r.key)
              }
              style={buttonStyle("danger", {
                fontSize: 12,
              })}
            >
              모두 되돌리기
            </button>
          )}
        </div>
        {changedCount === 0 ? (
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
            }}
          >
            아직 바꾼 과목이 없습니다. 위에서 과목을 누르면 선택/취소가 바뀝니다.
          </div>
        ) : (
          <div
            style={{
              fontSize: 12.5,
              lineHeight: 1.7,
            }}
          >
            {r.change.added.map((e, i) => (
              <div
                key={`a${i}`}
                style={{
                  color: OK,
                }}
              >
                {"+ \uCD94\uAC00 \u00B7 "}
                {e.name}
                {" ("}
                {formatSemester(e.semester)})
              </div>
            ))}
            {r.change.removed.map((e, i) => (
              <div
                key={`r${i}`}
                style={{
                  color: WARN,
                }}
              >
                {"\u2212 \uCDE8\uC18C \u00B7 "}
                {e.name}
                {" ("}
                {formatSemester(e.semester)})
              </div>
            ))}
          </div>
        )}
        <ChangeApplyBox r={r} setChangeApplied={setChangeApplied} />
        <div
          style={{
            fontSize: 12.5,
            fontWeight: 700,
            marginTop: 12,
            marginBottom: 4,
          }}
        >
          상담 메모
        </div>
        <textarea
          value={r.change.memo}
          onChange={(e) => setChangeMemo(r, e.target.value)}
          rows={3}
          placeholder="예: 9/15 상담 — 예술 학점 부족으로 3-2 미술과 매체 추가, 학생·학부모 동의"
          style={{
            ...inputStyle,
            width: "100%",
            fontSize: 13,
            padding: "8px 10px",
            resize: "vertical",
          }}
        />
        {r.change.updatedAt && (
          <div
            style={{
              fontSize: 11,
              color: MUTED,
              marginTop: 4,
            }}
          >
            {"\uB9C8\uC9C0\uB9C9 \uC218\uC815 "}
            {new Date(r.change.updatedAt).toLocaleString("ko-KR")}
          </div>
        )}
      </div>
    </div>
  );
}
function ChangePage({
  results,
  groups,
  reqOverride,
  priorityGroups,
  poolDefs,
  selectedKey,
  setSelectedKey,
  toggleCourseChange,
  setChangeMemo,
  resetCourseChange,
  onOpenCheck,
  onOpenTransfer,
  printOne,
  setChangeApplied,
}) {
  const [filterId, setFilterId] = useState("fail");
  const filters = [
    {
      id: "fail",
      label: "확인 필요",
      test: (r) => !r.pass,
    },
    {
      id: "changed",
      label: "상담·변경",
      test: (r) => r.hasChanges,
    },
    {
      id: "unapplied",
      label: "시스템 미반영",
      test: isChangeUnapplied,
    },
    {
      id: "all",
      label: "전체",
      test: () => true,
    },
  ];
  const edited = results.filter((r) => r.change.applyState !== "none");
  const appliedCount = edited.filter((r) => r.change.applyState === "applied").length;
  const selectedResult = results.find((r) => r.key === selectedKey) || null;
  return (
    <div>
      <PageIntro title="④ 선택과목 변경 (상담 후 수정)">
        ③에서 ‘확인 필요’로 나온 학생과 상담한 뒤 바뀐 선택과목을 여기서 고칩니다. 과목을 누르면 선택/취소가 바뀌고
        학점과 판정이 바로 다시 계산되어 ③ 점검 결과, 확인서, CSV에 반영됩니다. 원본 수강신청 파일은 바뀌지 않으므로{" "}
        <b
          style={{
            color: INK,
          }}
        >
          실제 수강신청 시스템(고교학점제 누리집·압핀)에도 같은 내용을 반영
        </b>
        한 뒤, 변경 내역 아래{" "}
        <b
          style={{
            color: INK,
          }}
        >
          ‘수강신청 시스템에도 이 변경을 반영했습니다’
        </b>
        를 체크하세요. 변경 내용과 반영 표시는 ‘진행 상황 저장’으로 파일에 남습니다.
      </PageIntro>
      {edited.length > 0 && (
        <div
          style={{
            display: "flex",
            gap: 8,
            flexWrap: "wrap",
            alignItems: "center",
            marginBottom: 14,
            fontSize: 12.5,
          }}
        >
          <span
            style={{
              background: "#fff",
              border: `1px solid ${LINE}`,
              borderRadius: 999,
              padding: "4px 11px",
              fontWeight: 700,
            }}
          >
            {"\uACFC\uBAA9\uC744 \uBC14\uAFBC \uD559\uC0DD "}
            {edited.length}명
          </span>
          <span
            style={{
              background: OK_BG,
              color: OK,
              borderRadius: 999,
              padding: "4px 11px",
              fontWeight: 700,
            }}
          >
            {"\uC2DC\uC2A4\uD15C \uBC18\uC601\uD568 "}
            {appliedCount}명
          </span>
          <button
            onClick={() => setFilterId("unapplied")}
            style={{
              ...chipStyle(filterId === "unapplied"),
              fontSize: 12.5,
              padding: "4px 11px",
              color: edited.length - appliedCount ? WARN : MUTED,
              borderColor: edited.length - appliedCount ? "#E8C9BB" : LINE,
              background: edited.length - appliedCount ? WARN_BG : "#fff",
            }}
          >
            {"\uC544\uC9C1 \uBC18\uC601 \uD45C\uC2DC \uC5C6\uC74C "}
            {edited.length - appliedCount}명 — 목록 보기
          </button>
        </div>
      )}
      {results.length === 0 ? (
        <EmptyPanel text="편제표와 수강신청 파일을 올리면 학생 목록이 나타납니다." />
      ) : (
        <div className="two-col">
          <StudentListPanel
            results={results}
            filters={filters}
            filterId={filterId}
            onFilter={setFilterId}
            selectedKey={selectedKey}
            onSelect={setSelectedKey}
            emptyText={
              filterId === "fail"
                ? "확인 필요 학생이 없습니다. 모두 충족입니다."
                : filterId === "unapplied"
                  ? "반영 표시가 필요한 학생이 없습니다."
                  : "해당하는 학생이 없습니다."
            }
          />
          {selectedResult ? (
            <ChangeEditor
              r={selectedResult}
              groups={groups}
              reqOverride={reqOverride}
              priorityGroups={priorityGroups}
              poolDefs={poolDefs}
              toggleCourseChange={toggleCourseChange}
              setChangeMemo={setChangeMemo}
              resetCourseChange={resetCourseChange}
              onOpenCheck={onOpenCheck}
              onOpenTransfer={onOpenTransfer}
              printOne={printOne}
              setChangeApplied={setChangeApplied}
            />
          ) : (
            <EmptyPanel text="왼쪽 목록에서 학생을 고르세요. 판정이 ‘확인필요’에서 ‘충족’으로 바뀌면 목록의 표시도 바로 바뀝니다." />
          )}
        </div>
      )}
    </div>
  );
}
// ================= PAGE 5: 전입생 =================
function TransferEditor({
  r,
  info,
  manual,
  courses,
  groups,
  reqOverride,
  priorityGroups,
  subjectLookup,
  updateTransfer,
  removeTransfer,
  updateManualStudent,
  removeManualStudent,
  onOpenCheck,
  onOpenChange,
  neisLoaded,
  neisCandidates,
  applyNeisRows,
}) {
  const prior = semestersBefore(info.firstSem);
  const rows = info.priorCourses || [];
  const setRows = (fn) =>
    updateTransfer(r.key, (cur) => ({
      priorCourses: fn(cur.priorCourses || []),
    }));
  const changeRow = (id, patch) =>
    setRows((list) =>
      list.map((x) =>
        x.id === id
          ? {
              ...x,
              ...patch,
            }
          : x,
      ),
    );
  const removeRow = (id) => setRows((list) => list.filter((x) => x.id !== id));
  const addRow = (sem) =>
    setRows((list) => [
      ...list,
      {
        id: newId("pc"),
        semester: sem,
        group: "",
        name: "",
        credit: "",
      },
    ]);
  const fillFromCurriculum = (sem) =>
    setRows((list) => {
      const have = new Set(list.filter((x) => x.semester === sem).map((x) => normalizeName(x.name)));
      const add = courses
        .filter((c) => c.semester === sem && c.gubun === "학교지정" && !have.has(c.key))
        .map((c) => ({
          id: newId("pc"),
          semester: sem,
          group: toCurriculumGroup(c.group),
          name: c.name,
          credit: c.credit,
        }));
      return [...list, ...add];
    });
  const onName = (row, value) => {
    const hit = subjectLookup.get(normalizeName(value));
    const patch = {
      name: value,
    };
    if (hit) {
      if (!row.group) patch.group = hit.group;
      if (row.credit === "" || row.credit == null || Number(row.credit) === 0) patch.credit = hit.credit;
    }
    changeRow(row.id, patch);
  };
  const ignored = rows.filter((x) => !prior.includes(x.semester));
  const designatedCount = (sem) => courses.filter((c) => c.semester === sem && c.gubun === "학교지정").length;
  return (
    <div
      style={{
        ...CARD,
        padding: 18,
        minWidth: 0,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          gap: 10,
          flexWrap: "wrap",
        }}
      >
        <div>
          <div
            style={{
              fontSize: 18,
              fontWeight: 800,
            }}
          >
            {r.name || "(이름 없음)"}
            <StudentTags r={r} />
          </div>
          <div
            style={{
              fontSize: 12.5,
              color: MUTED,
              marginTop: 2,
            }}
          >
            {r.studentId}
            {" \u00B7 "}
            {r.classLabel}
          </div>
        </div>
        <div
          style={{
            display: "flex",
            gap: 6,
            alignItems: "center",
            flexWrap: "wrap",
          }}
        >
          <PassBadge pass={r.pass} size={12.5} />
          <span
            style={{
              fontSize: 13,
              fontWeight: 700,
            }}
          >
            {"\uCD1D "}
            {r.total}학점
          </span>
          <button
            onClick={() => onOpenCheck(r.key)}
            style={buttonStyle("ghost", {
              fontSize: 12,
            })}
          >
            ③ 상세 보기
          </button>
          <button
            onClick={() => onOpenChange(r.key)}
            style={buttonStyle("ghost", {
              fontSize: 12,
            })}
          >
            ④ 선택과목
          </button>
          <button
            onClick={() =>
              window.confirm(
                `전입생 등록을 해제하면 입력한 전적교 과목 ${rows.length}개도 함께 지워집니다. 계속할까요?`,
              ) && removeTransfer(r.key)
            }
            style={buttonStyle("danger", {
              fontSize: 12,
            })}
          >
            등록 해제
          </button>
        </div>
      </div>
      {manual && (
        <div
          style={{
            marginTop: 12,
            background: PAPER,
            borderRadius: 8,
            padding: 10,
            display: "flex",
            gap: 6,
            flexWrap: "wrap",
            alignItems: "center",
            fontSize: 12.5,
          }}
        >
          <span
            style={{
              color: MUTED,
              fontWeight: 700,
            }}
          >
            직접 추가한 학생
          </span>
          <input
            value={manual.name}
            onChange={(e) =>
              updateManualStudent(manual.id, {
                name: e.target.value,
              })
            }
            placeholder="이름"
            style={{
              ...inputStyle,
              width: 96,
            }}
          />
          <select
            value={String(manual.grade || "")}
            onChange={(e) =>
              updateManualStudent(manual.id, {
                grade: e.target.value,
              })
            }
            style={{
              ...inputStyle,
              width: 70,
            }}
          >
            <option value="">학년</option>
            {["1", "2", "3"].map((g) => (
              <option key={g} value={g}>
                {g}학년
              </option>
            ))}
          </select>
          <input
            type="number"
            min="1"
            value={manual.classNum}
            onChange={(e) =>
              updateManualStudent(manual.id, {
                classNum: e.target.value,
              })
            }
            placeholder="반"
            style={{
              ...inputStyle,
              width: 56,
            }}
          />
          <input
            type="number"
            min="1"
            value={manual.number}
            onChange={(e) =>
              updateManualStudent(manual.id, {
                number: e.target.value,
              })
            }
            placeholder="번호"
            style={{
              ...inputStyle,
              width: 60,
            }}
          />
          <button
            onClick={() =>
              window.confirm(
                `${manual.name || "이 학생"}을(를) 목록에서 삭제할까요? 입력한 전입·변경·공동교육과정 내용도 함께 지워집니다.`,
              ) && removeManualStudent(manual.id)
            }
            style={buttonStyle("danger", {
              fontSize: 12,
              marginLeft: "auto",
            })}
          >
            학생 삭제
          </button>
        </div>
      )}
      <div
        style={{
          marginTop: 14,
          display: "grid",
          gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
          gap: 10,
        }}
      >
        <label
          style={{
            fontSize: 12,
            color: MUTED,
          }}
        >
          본교에서 처음 다닌 학기
          <select
            value={info.firstSem}
            onChange={(e) =>
              updateTransfer(r.key, {
                firstSem: e.target.value,
              })
            }
            style={{
              ...inputStyle,
              width: "100%",
              marginTop: 3,
              fontSize: 13,
            }}
          >
            {SEMESTER_CODES.slice(1).map((s) => (
              <option key={s} value={s}>
                {formatSemester(s)}부터 본교
              </option>
            ))}
          </select>
        </label>
        <label
          style={{
            fontSize: 12,
            color: MUTED,
          }}
        >
          전적교
          <input
            value={info.prevSchool}
            onChange={(e) =>
              updateTransfer(r.key, {
                prevSchool: e.target.value,
              })
            }
            placeholder="예: ○○고등학교"
            style={{
              ...inputStyle,
              width: "100%",
              marginTop: 3,
              fontSize: 13,
            }}
          />
        </label>
        <label
          style={{
            fontSize: 12,
            color: MUTED,
          }}
        >
          메모
          <input
            value={info.memo}
            onChange={(e) =>
              updateTransfer(r.key, {
                memo: e.target.value,
              })
            }
            placeholder="예: 2026.3 전입, 생기부 확인"
            style={{
              ...inputStyle,
              width: "100%",
              marginTop: 3,
              fontSize: 13,
            }}
          />
        </label>
      </div>
      <div
        style={{
          marginTop: 10,
          fontSize: 12,
          color: MUTED,
          lineHeight: 1.6,
          background: PAPER,
          borderRadius: 8,
          padding: "8px 10px",
        }}
      >
        {prior.map(formatSemester).join(" · ")}
        {"\uC740(\uB294) \uBCF8\uAD50 \uD3B8\uC81C\uD45C \uB300\uC2E0 \uC544\uB798 "}
        <b
          style={{
            color: INK,
          }}
        >
          전적교 이수 과목
        </b>
        {
          "\uC73C\uB85C \uACC4\uC0B0\uD558\uACE0, \uBCF8\uAD50 \uD559\uAD50\uC9C0\uC815 \uACFC\uBAA9\uACFC \uC218\uAC15\uC2E0\uCCAD \uAE30\uB85D\uC740 "
        }
        {formatSemester(info.firstSem)}부터만 반영합니다. 전적교 과목도 2022 개정 과목명이면 이름을 넣을 때 교과(군)와
        기준학점이 자동으로 채워지니, 학교생활기록부를 보고 실제 이수 학점으로 고쳐주세요.
      </div>
      {neisLoaded && (
        <NeisImportBox
          info={info}
          candidates={neisCandidates || []}
          groups={groups}
          subjectLookup={subjectLookup}
          onApply={(plan) => applyNeisRows(r.key, plan.rows, plan.semesters)}
        />
      )}
      {prior.map((sem) => {
        const semRows = rows.filter((x) => x.semester === sem);
        const semTotal = semRows.reduce((a, x) => a + (Number(x.credit) || 0), 0);
        return (
          <div
            key={sem}
            style={{
              marginTop: 14,
              border: `1px solid ${LINE}`,
              borderRadius: 9,
              padding: 12,
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 8,
                flexWrap: "wrap",
                marginBottom: 8,
              }}
            >
              <span
                style={{
                  fontSize: 13.5,
                  fontWeight: 800,
                }}
              >
                {formatSemester(sem)}{" "}
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    color: MUTED,
                  }}
                >
                  {"\u00B7 \uC804\uC801\uAD50 \uC774\uC218 "}
                  {semRows.length}
                  {"\uACFC\uBAA9 "}
                  {semTotal}학점
                </span>
              </span>
              <span
                style={{
                  display: "flex",
                  gap: 6,
                }}
              >
                <button
                  onClick={() => addRow(sem)}
                  style={buttonStyle("soft", {
                    fontSize: 12,
                  })}
                >
                  + 과목 추가
                </button>
                {designatedCount(sem) > 0 && (
                  <button
                    onClick={() => fillFromCurriculum(sem)}
                    style={buttonStyle("ghost", {
                      fontSize: 12,
                    })}
                    title="본교 편제표의 이 학기 학교지정 과목을 불러온 뒤, 전적교 학점에 맞게 고치세요"
                  >
                    본교 학교지정 과목 불러오기 ({designatedCount(sem)})
                  </button>
                )}
              </span>
            </div>
            {semRows.length === 0 ? (
              <div
                style={{
                  fontSize: 12.5,
                  color: WARN,
                }}
              >
                아직 입력한 과목이 없습니다. 이 학기는 0학점으로 계산됩니다.
              </div>
            ) : (
              <div
                style={{
                  overflowX: "auto",
                }}
              >
                <table
                  style={{
                    borderCollapse: "collapse",
                    width: "100%",
                    minWidth: 480,
                  }}
                >
                  <thead>
                    <tr>
                      <th
                        style={{
                          ...thStyle,
                          width: 150,
                        }}
                      >
                        교과(군)
                      </th>
                      <th style={thStyle}>과목명</th>
                      <th
                        style={{
                          ...thStyle,
                          width: 70,
                        }}
                      >
                        학점
                      </th>
                      <th
                        style={{
                          ...thStyle,
                          width: 36,
                        }}
                      />
                    </tr>
                  </thead>
                  <tbody>
                    {semRows.map((row) => (
                      <tr
                        key={row.id}
                        style={{
                          borderTop: `1px solid ${LINE}`,
                        }}
                      >
                        <td
                          style={{
                            padding: "4px 6px",
                          }}
                        >
                          <GroupSelect
                            value={row.group}
                            groups={groups}
                            onChange={(v) =>
                              changeRow(row.id, {
                                group: v,
                              })
                            }
                            style={{
                              width: "100%",
                              borderColor: row.group ? LINE : WARN,
                            }}
                          />
                        </td>
                        <td
                          style={{
                            padding: "4px 6px",
                          }}
                        >
                          <input
                            list="subject-names-transfer"
                            value={row.name}
                            onChange={(e) => onName(row, e.target.value)}
                            placeholder="과목명"
                            style={{
                              ...inputStyle,
                              width: "100%",
                              borderColor: row.name ? LINE : WARN,
                            }}
                          />
                        </td>
                        <td
                          style={{
                            padding: "4px 6px",
                          }}
                        >
                          <input
                            type="number"
                            min="0"
                            step="0.5"
                            value={row.credit}
                            onChange={(e) =>
                              changeRow(row.id, {
                                credit: e.target.value === "" ? "" : Number(e.target.value),
                              })
                            }
                            style={{
                              ...inputStyle,
                              width: "100%",
                              borderColor: Number(row.credit) > 0 ? LINE : WARN,
                            }}
                          />
                        </td>
                        <td
                          style={{
                            padding: "4px 6px",
                            textAlign: "center",
                          }}
                        >
                          <button
                            onClick={() => removeRow(row.id)}
                            style={{
                              border: "none",
                              background: "none",
                              cursor: "pointer",
                              color: MUTED,
                            }}
                            title="삭제"
                          >
                            <X size={13} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
      {ignored.length > 0 && (
        <div
          style={{
            marginTop: 14,
            background: WARN_BG,
            borderRadius: 8,
            padding: "10px 12px",
            fontSize: 12.5,
            color: WARN,
          }}
        >
          <div
            style={{
              fontWeight: 700,
              marginBottom: 4,
            }}
          >
            {
              "\uBCF8\uAD50 \uC7AC\uD559 \uD559\uAE30\uB85C \uC785\uB825\uB418\uC5B4 \uACC4\uC0B0\uC5D0\uC11C \uBE80 \uC804\uC801\uAD50 \uACFC\uBAA9 "
            }
            {ignored.length}개
          </div>
          {ignored.map((x) => (
            <div
              key={x.id}
              style={{
                display: "flex",
                justifyContent: "space-between",
                gap: 8,
              }}
            >
              <span>
                {formatSemester(x.semester)}
                {" \u00B7 "}
                {x.name || "(과목명 없음)"}
                {" \u00B7 "}
                {x.credit || 0}학점
              </span>
              <button
                onClick={() => removeRow(x.id)}
                style={{
                  border: "none",
                  background: "none",
                  color: WARN,
                  cursor: "pointer",
                  fontSize: 12,
                }}
              >
                삭제
              </button>
            </div>
          ))}
          <div
            style={{
              fontSize: 11.5,
              marginTop: 4,
              opacity: 0.85,
            }}
          >
            ‘본교에서 처음 다닌 학기’를 바꿨다면 확인해주세요.
          </div>
        </div>
      )}
      {r.excludedBeforeTransfer.length > 0 && (
        <div
          style={{
            marginTop: 10,
            fontSize: 12,
            color: MUTED,
          }}
        >
          {"\uC804\uC785 \uC774\uC804 \uD559\uAE30\uC758 \uC218\uAC15\uC2E0\uCCAD \uAE30\uB85D "}
          {r.excludedBeforeTransfer.length}과목(
          {r.excludedBeforeTransfer.map((e) => `${e.name} ${formatSemester(e.semester)}`).join(", ")})은 계산에서
          뺐습니다.
        </div>
      )}
      <div
        style={{
          marginTop: 18,
          borderTop: `1px solid ${LINE}`,
          paddingTop: 14,
        }}
      >
        <div
          style={{
            fontSize: 13.5,
            fontWeight: 700,
            marginBottom: 8,
          }}
        >
          이 학생의 현재 판정
        </div>
        {r.reasons.length > 0 ? (
          <ReasonList reasons={r.reasons} />
        ) : (
          <div
            style={{
              fontSize: 12.5,
              color: OK,
              fontWeight: 700,
            }}
          >
            ✓ 확인이 필요한 항목이 없습니다.
          </div>
        )}
        <div
          style={{
            marginTop: 10,
          }}
        >
          <GroupCreditGrid result={r} groups={groups} reqOverride={reqOverride} priorityGroups={priorityGroups} />
        </div>
      </div>
    </div>
  );
}
// 선택한 전입생을 올린 나이스 자료에서 찾아, 본교에 오기 전 학기 과목을 한 번에 채웁니다.
function NeisImportBox({ info, candidates, groups, subjectLookup, onApply }) {
  if (candidates.length === 0) {
    return (
      <div
        style={{
          marginTop: 12,
          fontSize: 12.5,
          color: MUTED,
          background: PAPER,
          borderRadius: 8,
          padding: "9px 11px",
          lineHeight: 1.55,
        }}
      >
        올린 나이스 교과학습발달상황 파일에서 이 학생을 찾지 못했습니다(이름으로 찾습니다). 이 학생이 있는 반의 파일을
        올렸는지 확인하세요.
      </div>
    );
  }
  return (
    <div
      style={{
        marginTop: 12,
        display: "flex",
        flexDirection: "column",
        gap: 8,
      }}
    >
      {candidates.map((c) => {
        const plan = neisPriorRows(c.student, info.firstSem, groups, subjectLookup);
        const existing = (info.priorCourses || []).filter(
          (p) => plan.semesters.includes(p.semester) && String(p.name || "").trim(),
        ).length;
        const noGroup = plan.rows.filter((x) => !x.group).length;
        const skippedSems = Array.from(new Set(plan.skipped.map((x) => x.semester))).sort(compareSemester);
        const apply = () => {
          if (
            existing > 0 &&
            !window.confirm(
              `${plan.semesters.map(formatSemester).join(" · ")}에 이미 입력한 전적교 과목 ${existing}개를 나이스 자료의 ${plan.rows.length}과목으로 바꿉니다. 계속할까요?`,
            )
          )
            return;
          onApply(plan);
        };
        return (
          <div
            key={c.fileId + ":" + c.index}
            style={{
              background: ACCENT_BG,
              borderRadius: 9,
              padding: "10px 12px",
              fontSize: 12.5,
              lineHeight: 1.6,
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                gap: 8,
                flexWrap: "wrap",
              }}
            >
              <span>
                <b
                  style={{
                    color: ACCENT,
                  }}
                >
                  나이스 교과학습발달상황
                </b>
                {" \u00B7 "}
                {neisPlaceText(c)} {c.student.name}
                {!c.exact && candidates.length > 1 && (
                  <span
                    style={{
                      color: WARN,
                      fontWeight: 700,
                    }}
                  >
                    {" \u00B7 \uC774\uB984\uB9CC \uAC19\uC740 \uD559\uC0DD\uC774 "}
                    {candidates.length}명 — 반·번호를 확인하세요
                  </span>
                )}
              </span>
              <button
                onClick={apply}
                disabled={plan.rows.length === 0}
                style={buttonStyle(plan.rows.length > 0 ? "primary" : "disabled", {
                  fontSize: 12,
                })}
              >
                이 과목으로 채우기
              </button>
            </div>
            <div
              style={{
                color: MUTED,
              }}
            >
              {plan.rows.length > 0
                ? `${plan.semesters.map(formatSemester).join(" · ")} 과목 ${plan.rows.length}개(${plan.credit}학점)를 전적교 이수 과목으로 가져옵니다.`
                : `${formatSemester(info.firstSem)} 이전 학기의 과목이 이 자료에 없습니다.`}
              {plan.skipped.length > 0 &&
                ` 본교 재학 학기(${skippedSems.map(formatSemester).join(" · ")}) 과목 ${plan.skipped.length}개는 가져오지 않습니다.`}
              {existing > 0 && ` 같은 학기에 이미 입력한 과목 ${existing}개는 바뀝니다.`}
              {noGroup > 0 && (
                <span
                  style={{
                    color: WARN,
                  }}
                >
                  {" \uAD50\uACFC(\uAD70)\uC744 \uC54C\uC544\uBCF4\uC9C0 \uBABB\uD55C \uACFC\uBAA9 "}
                  {noGroup}개는 가져온 뒤 직접 골라주세요.
                </span>
              )}
            </div>
          </div>
        );
      })}
    </div>
  );
}
// 나이스 교과학습발달상황 엑셀 올리기 · 등록된 전입생 한꺼번에 채우기
function NeisUploadCard({
  neisFiles,
  handleNeisFiles,
  removeNeisFile,
  transferResults,
  transferInfo,
  groups,
  subjectLookup,
  applyNeisRows,
}) {
  const okFiles = neisFiles.filter((f) => f.status === "ok");
  const warnings = okFiles.flatMap((f) => f.warnings || []);
  const plans = okFiles.length
    ? transferResults.map((r) => {
        const cands = findNeisCandidates(neisFiles, r);
        const info = transferInfo[r.key];
        const plan = cands.length === 1 ? neisPriorRows(cands[0].student, info.firstSem, groups, subjectLookup) : null;
        const hasInput = plan
          ? (info.priorCourses || []).some((p) => plan.semesters.includes(p.semester) && String(p.name || "").trim())
          : false;
        return {
          r,
          cands,
          plan,
          fillable: !!(plan && plan.rows.length > 0 && !hasInput),
        };
      })
    : [];
  const found = plans.filter((p) => p.cands.length > 0).length;
  const fillable = plans.filter((p) => p.fillable);
  const fillAll = () => {
    if (
      !window.confirm(
        `전입생 ${fillable.length}명의 전적교 과목을 나이스 자료로 채웁니다. 이미 과목을 입력한 학기가 있는 학생은 건너뜁니다. 계속할까요?`,
      )
    )
      return;
    fillable.forEach((p) => applyNeisRows(p.r.key, p.plan.rows, p.plan.semesters));
  };
  return (
    <div
      style={{
        ...CARD,
        padding: 12,
      }}
    >
      <div
        style={{
          fontSize: 13,
          fontWeight: 700,
          marginBottom: 4,
        }}
      >
        나이스 교과학습발달상황 (엑셀)
      </div>
      <div
        style={{
          fontSize: 11.5,
          color: MUTED,
          lineHeight: 1.55,
          marginBottom: 8,
        }}
      >
        {
          "\uB098\uC774\uC2A4 [\uD559\uAD50\uC0DD\uD65C\uAE30\uB85D\uBD80 \u2192 \uD559\uC0DD\uBD80 \uD56D\uBAA9\uBCC4 \uC870\uD68C \u2192 \uAD50\uACFC\uD559\uC2B5\uBC1C\uB2EC\uC0C1\uD669]\uC5D0\uC11C \uC804\uC785\uC0DD\uC774 \uC788\uB294 \uBC18\uC744 "
        }
        <b
          style={{
            color: INK,
          }}
        >
          엑셀
        </b>
        로 내려받아 올리면 전적교 과목을 한 번에 채울 수 있습니다. 교과·과목·학점만 읽고 성적은 읽지 않습니다.
      </div>
      <RegFileDropzone onFiles={handleNeisFiles} label="나이스 엑셀을 끌어다 놓거나" />
      {neisFiles.length > 0 && (
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 4,
            marginTop: 8,
          }}
        >
          {neisFiles.map((f) => (
            <div
              key={f.id}
              style={{
                display: "flex",
                alignItems: "flex-start",
                justifyContent: "space-between",
                gap: 6,
                fontSize: 12,
                background: f.status === "error" ? WARN_BG : PAPER,
                borderRadius: 7,
                padding: "6px 8px",
              }}
            >
              <div
                style={{
                  display: "flex",
                  gap: 6,
                  minWidth: 0,
                }}
              >
                {f.status === "error" ? (
                  <AlertCircle
                    size={13}
                    color={WARN}
                    style={{
                      flexShrink: 0,
                      marginTop: 2,
                    }}
                  />
                ) : (
                  <CheckCircle2
                    size={13}
                    color={OK}
                    style={{
                      flexShrink: 0,
                      marginTop: 2,
                    }}
                  />
                )}
                <span
                  style={{
                    minWidth: 0,
                    wordBreak: "break-all",
                  }}
                >
                  {f.name}
                  {f.status === "ok" ? (
                    <span
                      style={{
                        color: MUTED,
                      }}
                    >
                      {" "}
                      {"\u00B7 "}
                      {f.classInfo ? `${f.classInfo.grade}학년 ${f.classInfo.classNum}반` : "반 정보 없음"}
                      {" \u00B7 \uD559\uC0DD "}
                      {f.students.length}명
                    </span>
                  ) : (
                    <span
                      style={{
                        color: WARN,
                      }}
                    >
                      {" \u00B7 "}
                      {f.error}
                    </span>
                  )}
                </span>
              </div>
              <button
                onClick={() => removeNeisFile(f.id)}
                title="목록에서 빼기"
                style={{
                  border: "none",
                  background: "none",
                  cursor: "pointer",
                  color: MUTED,
                  flexShrink: 0,
                }}
              >
                <X size={13} />
              </button>
            </div>
          ))}
        </div>
      )}
      {warnings.length > 0 && (
        <details
          style={{
            marginTop: 6,
            fontSize: 11.5,
            color: WARN,
          }}
        >
          <summary
            style={{
              cursor: "pointer",
            }}
          >
            {"\uD655\uC778\uD560 \uC810 "}
            {warnings.length}건
          </summary>
          {warnings.map((w, i) => (
            <div key={i}>
              {"\u00B7 "}
              {w}
            </div>
          ))}
        </details>
      )}
      {okFiles.length > 0 && transferResults.length > 0 && (
        <div
          style={{
            marginTop: 8,
            fontSize: 12,
            lineHeight: 1.55,
          }}
        >
          {"\uB4F1\uB85D\uB41C \uC804\uC785\uC0DD "}
          {transferResults.length}
          {"\uBA85 \uC911 "}
          <b>{found}명</b>을 나이스 자료에서 찾았습니다.
          {fillable.length > 0 && (
            <button
              onClick={fillAll}
              style={buttonStyle("soft", {
                fontSize: 12,
                marginTop: 6,
                width: "100%",
                justifyContent: "center",
              })}
            >
              {"\uC544\uC9C1 \uBE44\uC5B4 \uC788\uB294 "}
              {fillable.length}명 모두 채우기
            </button>
          )}
        </div>
      )}
      {okFiles.length > 0 && transferResults.length === 0 && (
        <div
          style={{
            marginTop: 8,
            fontSize: 12,
            color: MUTED,
          }}
        >
          위에서 전입생을 먼저 등록하면, 학생별로 나이스 과목을 채울 수 있습니다.
        </div>
      )}
      <div
        style={{
          marginTop: 8,
          fontSize: 11,
          color: MUTED,
          lineHeight: 1.5,
        }}
      >
        올린 나이스 파일은 진행 상황 저장 파일에 들어가지 않습니다(채운 과목만 저장됩니다).
      </div>
    </div>
  );
}
function TransferPage({
  results,
  transferInfo,
  courses,
  groups,
  reqOverride,
  priorityGroups,
  selectedKey,
  setSelectedKey,
  registerTransfer,
  updateTransfer,
  removeTransfer,
  manualStudents,
  addManualStudent,
  updateManualStudent,
  removeManualStudent,
  onOpenCheck,
  onOpenChange,
  neisFiles,
  handleNeisFiles,
  removeNeisFile,
  applyNeisRows,
}) {
  const [q, setQ] = useState("");
  const [showManual, setShowManual] = useState(false);
  const [form, setForm] = useState({
    name: "",
    grade: "2",
    classNum: "",
    number: "",
  });
  const subjectLookup = useMemo(() => buildSubjectLookup(courses), [courses]);
  if (courses.length === 0) return <EmptyPanel text="편제표를 먼저 올려주세요." />;
  const transferResults = results.filter((r) => transferInfo[r.key]).sort(byClassThenId);
  const selectedResult = results.find((r) => r.key === selectedKey && transferInfo[r.key]) || null;
  const query = q.trim();
  const candidates = query
    ? results
        .filter((r) => !transferInfo[r.key] && (r.name.includes(query) || r.studentId.includes(query)))
        .sort(byClassThenId)
        .slice(0, 8)
    : [];
  const manualByKey = new Map(manualStudents.map((m) => ["manual:" + m.id, m]));
  return (
    <div>
      <PageIntro title="⑤ 전입생 — 학생별 전입 이전 이수 과목">
        {
          "\uC804\uC785\uC0DD\uC740 \uC804\uC785 \uC804 \uD559\uAE30\uB97C \uB2E4\uB978 \uD559\uAD50 \uD3B8\uC81C\uD45C\uB85C \uC774\uC218\uD574 \uACFC\uBAA9\uACFC \uD559\uC810\uC774 \uBCF8\uAD50\uC640 \uB2E4\uB985\uB2C8\uB2E4. \uD559\uC0DD\uB9C8\uB2E4 "
        }
        <b
          style={{
            color: INK,
          }}
        >
          본교에서 처음 다닌 학기
        </b>
        {
          "\uB97C \uC815\uD558\uBA74, \uADF8 \uC774\uC804 \uD559\uAE30\uB294 \uBCF8\uAD50 \uD3B8\uC81C\uD45C \uB300\uC2E0 \uC5EC\uAE30\uC5D0 \uC785\uB825\uD55C "
        }
        <b
          style={{
            color: INK,
          }}
        >
          전적교 이수 과목
        </b>
        으로 계산하고, 본교 학교지정 과목과 수강신청 기록은 전입 학기부터만 반영합니다. 과목과 학점은
        학교생활기록부(전적교 성적)를 보고 입력하세요.
      </PageIntro>
      <SubjectDatalist id="subject-names-transfer" lookup={subjectLookup} />
      <div className="two-col">
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            gap: 12,
            minWidth: 0,
          }}
        >
          <div
            style={{
              ...CARD,
              padding: 12,
            }}
          >
            <div
              style={{
                fontSize: 13,
                fontWeight: 700,
                marginBottom: 8,
              }}
            >
              전입생 등록
            </div>
            <input
              value={q}
              onChange={(e) => setQ(e.target.value)}
              placeholder="이름·학번으로 학생 찾기"
              style={{
                ...inputStyle,
                width: "100%",
              }}
            />
            {candidates.length > 0 && (
              <div
                style={{
                  display: "flex",
                  flexDirection: "column",
                  gap: 2,
                  marginTop: 6,
                }}
              >
                {candidates.map((r) => (
                  <button
                    key={r.key}
                    onClick={() => {
                      registerTransfer(r);
                      setQ("");
                    }}
                    style={{
                      textAlign: "left",
                      border: `1px solid ${LINE}`,
                      borderRadius: 7,
                      padding: "6px 9px",
                      cursor: "pointer",
                      background: "#fff",
                      fontFamily: FONT,
                      fontSize: 12.5,
                    }}
                  >
                    {"+ "}
                    <b>{r.name}</b>{" "}
                    <span
                      style={{
                        color: MUTED,
                      }}
                    >
                      {r.studentId}
                      {" \u00B7 "}
                      {r.classLabel}
                    </span>
                  </button>
                ))}
              </div>
            )}
            {query && candidates.length === 0 && (
              <div
                style={{
                  fontSize: 12,
                  color: MUTED,
                  marginTop: 6,
                  lineHeight: 1.5,
                }}
              >
                일치하는 학생이 없습니다(이미 등록된 학생 제외). 수강신청 파일에 없는 학생이면 아래에서 직접 추가하세요.
              </div>
            )}
            <button
              onClick={() => setShowManual((v) => !v)}
              style={buttonStyle("ghost", {
                fontSize: 12,
                marginTop: 10,
                width: "100%",
                justifyContent: "center",
              })}
            >
              {showManual ? "닫기" : "수강신청 명단에 없는 학생 직접 추가"}
            </button>
            {showManual && (
              <div
                style={{
                  marginTop: 8,
                  display: "flex",
                  flexDirection: "column",
                  gap: 6,
                }}
              >
                <input
                  value={form.name}
                  onChange={(e) =>
                    setForm((f) => ({
                      ...f,
                      name: e.target.value,
                    }))
                  }
                  placeholder="이름"
                  style={{
                    ...inputStyle,
                    width: "100%",
                  }}
                />
                <div
                  style={{
                    display: "flex",
                    gap: 6,
                  }}
                >
                  <select
                    value={form.grade}
                    onChange={(e) =>
                      setForm((f) => ({
                        ...f,
                        grade: e.target.value,
                      }))
                    }
                    style={{
                      ...inputStyle,
                      flex: 1,
                    }}
                  >
                    {["1", "2", "3"].map((g) => (
                      <option key={g} value={g}>
                        {g}학년
                      </option>
                    ))}
                  </select>
                  <input
                    type="number"
                    min="1"
                    value={form.classNum}
                    onChange={(e) =>
                      setForm((f) => ({
                        ...f,
                        classNum: e.target.value,
                      }))
                    }
                    placeholder="반"
                    style={{
                      ...inputStyle,
                      width: 60,
                    }}
                  />
                  <input
                    type="number"
                    min="1"
                    value={form.number}
                    onChange={(e) =>
                      setForm((f) => ({
                        ...f,
                        number: e.target.value,
                      }))
                    }
                    placeholder="번호"
                    style={{
                      ...inputStyle,
                      width: 64,
                    }}
                  />
                </div>
                <button
                  disabled={!form.name.trim()}
                  onClick={() => {
                    addManualStudent(form);
                    setForm({
                      name: "",
                      grade: form.grade,
                      classNum: "",
                      number: "",
                    });
                    setShowManual(false);
                  }}
                  style={buttonStyle(form.name.trim() ? "primary" : "disabled", {
                    justifyContent: "center",
                  })}
                >
                  추가
                </button>
                <div
                  style={{
                    fontSize: 11.5,
                    color: MUTED,
                    lineHeight: 1.5,
                  }}
                >
                  현재 학년의 반·번호를 넣으면 ③ 학급별 보기에 들어갑니다. 본교에서 이수한 선택과목은 ④ 선택과목
                  변경에서 입력하세요.
                </div>
              </div>
            )}
          </div>
          <NeisUploadCard
            neisFiles={neisFiles}
            handleNeisFiles={handleNeisFiles}
            removeNeisFile={removeNeisFile}
            transferResults={transferResults}
            transferInfo={transferInfo}
            groups={groups}
            subjectLookup={subjectLookup}
            applyNeisRows={applyNeisRows}
          />
          <div
            style={{
              ...CARD,
              padding: 12,
            }}
          >
            <div
              style={{
                fontSize: 13,
                fontWeight: 700,
                marginBottom: 6,
              }}
            >
              {"\uB4F1\uB85D\uB41C \uC804\uC785\uC0DD "}
              {transferResults.length}명
            </div>
            {transferResults.length === 0 && (
              <div
                style={{
                  fontSize: 12.5,
                  color: MUTED,
                }}
              >
                아직 없습니다.
              </div>
            )}
            {transferResults.map((r) => (
              <button
                key={r.key}
                onClick={() => setSelectedKey(r.key)}
                style={{
                  width: "100%",
                  textAlign: "left",
                  border: "none",
                  borderRadius: 7,
                  padding: "7px 9px",
                  cursor: "pointer",
                  background: r.key === selectedKey ? ACCENT_BG : "transparent",
                  fontFamily: FONT,
                }}
              >
                <div
                  style={{
                    display: "flex",
                    justifyContent: "space-between",
                    alignItems: "center",
                    gap: 6,
                  }}
                >
                  <span
                    style={{
                      fontSize: 13,
                      fontWeight: 700,
                      color: INK,
                    }}
                  >
                    {r.name || "(이름 없음)"}
                  </span>
                  <PassBadge pass={r.pass} size={10.5} />
                </div>
                <div
                  style={{
                    fontSize: 11.5,
                    color: MUTED,
                  }}
                >
                  {r.studentId}
                  {" \u00B7 "}
                  {formatSemester(r.firstSem)}
                  {"\uBD80\uD130 \uBCF8\uAD50 \u00B7 \uC804\uC801\uAD50 \uACFC\uBAA9 "}
                  {r.priorCourses.length}개
                </div>
              </button>
            ))}
          </div>
        </div>
        {selectedResult ? (
          <TransferEditor
            r={selectedResult}
            info={transferInfo[selectedResult.key]}
            manual={manualByKey.get(selectedResult.key) || null}
            courses={courses}
            groups={groups}
            reqOverride={reqOverride}
            priorityGroups={priorityGroups}
            subjectLookup={subjectLookup}
            updateTransfer={updateTransfer}
            removeTransfer={removeTransfer}
            updateManualStudent={updateManualStudent}
            removeManualStudent={removeManualStudent}
            onOpenCheck={onOpenCheck}
            onOpenChange={onOpenChange}
            neisLoaded={neisFiles.some((f) => f.status === "ok")}
            neisCandidates={findNeisCandidates(neisFiles, selectedResult)}
            applyNeisRows={applyNeisRows}
          />
        ) : (
          <EmptyPanel text="왼쪽에서 학생을 찾아 전입생으로 등록하거나, 등록된 전입생을 고르세요." />
        )}
      </div>
    </div>
  );
}
// ================= PAGE 6: 공동교육과정 · 비고 =================
function ExtraCourseCard({
  x,
  results,
  allByKey,
  groups,
  subjectLookup,
  updateExtraCourse,
  removeExtraCourse,
  addStudentToExtra,
  removeStudentFromExtra,
  onOpenCheck,
}) {
  const incomplete = !String(x.name || "").trim() || !(Number(x.credit) > 0) || !x.group;
  const excludeKeys = new Set(x.studentKeys);
  const onName = (value) => {
    const hit = subjectLookup.get(normalizeName(value));
    const patch = {
      name: value,
    };
    if (hit) {
      if (!x.group) patch.group = hit.group;
      if (x.credit === "" || x.credit == null || Number(x.credit) === 0) patch.credit = hit.credit;
    }
    updateExtraCourse(x.id, patch);
  };
  return (
    <div
      style={{
        ...CARD,
        padding: 14,
        marginBottom: 12,
        borderLeft: `4px solid ${incomplete ? WARN : ACCENT}`,
      }}
    >
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "130px minmax(160px, 2fr) 150px 80px 140px auto",
          gap: 6,
          alignItems: "center",
        }}
        className="extra-grid"
      >
        <select
          value={x.kind}
          onChange={(e) =>
            updateExtraCourse(x.id, {
              kind: e.target.value,
            })
          }
          style={inputStyle}
        >
          {EXTRA_KINDS.map((k) => (
            <option key={k} value={k}>
              {k}
            </option>
          ))}
        </select>
        <input
          list="subject-names-extra"
          value={x.name}
          onChange={(e) => onName(e.target.value)}
          placeholder="과목명"
          style={{
            ...inputStyle,
            borderColor: x.name ? LINE : WARN,
          }}
        />
        <GroupSelect
          value={x.group}
          groups={groups}
          onChange={(v) =>
            updateExtraCourse(x.id, {
              group: v,
            })
          }
          style={{
            borderColor: x.group ? LINE : WARN,
          }}
        />
        <input
          type="number"
          min="0"
          step="0.5"
          value={x.credit}
          onChange={(e) =>
            updateExtraCourse(x.id, {
              credit: e.target.value === "" ? "" : Number(e.target.value),
            })
          }
          placeholder="학점"
          style={{
            ...inputStyle,
            borderColor: Number(x.credit) > 0 ? LINE : WARN,
          }}
        />
        <select
          value={x.semester}
          onChange={(e) =>
            updateExtraCourse(x.id, {
              semester: e.target.value,
            })
          }
          style={inputStyle}
        >
          <option value="">학기 미지정</option>
          {SEMESTER_CODES.map((s) => (
            <option key={s} value={s}>
              {formatSemester(s)}
            </option>
          ))}
        </select>
        <button
          onClick={() =>
            (x.studentKeys.length === 0 ||
              window.confirm(`‘${x.name || "이 과목"}’과 수강 학생 ${x.studentKeys.length}명 연결을 삭제할까요?`)) &&
            removeExtraCourse(x.id)
          }
          style={buttonStyle("danger", {
            fontSize: 12,
            justifyContent: "center",
          })}
        >
          삭제
        </button>
      </div>
      <input
        value={x.memo}
        onChange={(e) =>
          updateExtraCourse(x.id, {
            memo: e.target.value,
          })
        }
        placeholder="비고 (예: 강원 온라인학교 2026-2, 거점학교 ○○고, 이수 확인서 제출)"
        style={{
          ...inputStyle,
          width: "100%",
          marginTop: 6,
        }}
      />
      {incomplete && (
        <div
          style={{
            fontSize: 12,
            color: WARN,
            marginTop: 6,
          }}
        >
          과목명·교과(군)·학점을 모두 입력해야 정확히 반영됩니다. (교과군이 없으면 총학점에만 들어갑니다)
        </div>
      )}
      <div
        style={{
          marginTop: 10,
          fontSize: 12.5,
          fontWeight: 700,
        }}
      >
        {"\uC218\uAC15 \uD559\uC0DD "}
        {x.studentKeys.length}명
      </div>
      <div
        style={{
          display: "flex",
          flexWrap: "wrap",
          gap: 5,
          marginTop: 6,
        }}
      >
        {x.studentKeys.length === 0 && (
          <span
            style={{
              fontSize: 12,
              color: MUTED,
            }}
          >
            아직 없습니다. 아래에서 학생을 추가하세요.
          </span>
        )}
        {x.studentKeys.map((k) => {
          const r = allByKey.get(k);
          const label = r
            ? `${r.name} ${r.studentId}`
            : `${(x.studentNames || {})[k] || nameFromStudentKey(k) || "?"} (연결 끊김)`;
          return (
            <span
              key={k}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: 4,
                fontSize: 12,
                background: r ? ACCENT_BG : WARN_BG,
                color: r ? ACCENT : WARN,
                borderRadius: 999,
                padding: "2px 4px 2px 9px",
              }}
            >
              <span
                onClick={() => r && onOpenCheck(k)}
                style={{
                  cursor: r ? "pointer" : "default",
                }}
              >
                {label}
              </span>
              <button
                onClick={() => removeStudentFromExtra(x.id, k)}
                style={{
                  border: "none",
                  background: "none",
                  cursor: "pointer",
                  color: "inherit",
                  padding: "0 3px",
                }}
                title="이 학생 빼기"
              >
                <X size={11} />
              </button>
            </span>
          );
        })}
      </div>
      <StudentAdder results={results} excludeKeys={excludeKeys} onAdd={(r) => addStudentToExtra(x.id, r)} />
    </div>
  );
}
function ExtraPage({
  results,
  allResults,
  extraCourses,
  groups,
  courses,
  addExtraCourse,
  updateExtraCourse,
  removeExtraCourse,
  addStudentToExtra,
  removeStudentFromExtra,
  onOpenCheck,
}) {
  const subjectLookup = useMemo(() => buildSubjectLookup(courses), [courses]);
  const allByKey = useMemo(() => new Map((allResults || results).map((r) => [r.key, r])), [allResults, results]);
  if (courses.length === 0) return <EmptyPanel text="편제표를 먼저 올려주세요." />;
  const withExtras = results.filter((r) => r.extraEntries.length > 0).sort(byClassThenId);
  return (
    <div>
      <PageIntro title="⑥ 공동교육과정 · 비고">
        {
          "\uACF5\uB3D9\uAD50\uC721\uACFC\uC815(\uC628\u00B7\uC624\uD504\uB77C\uC778), \uC628\uB77C\uC778\uD559\uAD50, \uD559\uAD50 \uBC16 \uAD50\uC721\uCC98\uB7FC \uBCF8\uAD50 \uC218\uAC15\uC2E0\uCCAD \uD30C\uC77C\uC5D0 \uC5C6\uB294 \uC774\uC218 \uACFC\uBAA9\uC744 "
        }
        <b
          style={{
            color: INK,
          }}
        >
          과목 단위
        </b>
        로 입력하고 들은 학생을 고릅니다. 같은 과목을 여러 학생이 들었다면 한 번만 입력하면 됩니다. 학점은 해당
        교과(군)와 총학점에 바로 반영되고 ③ 점검, 확인서, CSV에 함께 나옵니다. (택N 선택 개수에는 들어가지 않습니다.)
      </PageIntro>
      <SubjectDatalist id="subject-names-extra" lookup={subjectLookup} />
      <style>{`@media (max-width: 900px) { .extra-grid { grid-template-columns: 1fr 1fr !important; } }`}</style>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          marginBottom: 12,
          gap: 8,
          flexWrap: "wrap",
        }}
      >
        <div
          style={{
            fontSize: 13,
            color: MUTED,
          }}
        >
          {"\uC785\uB825\uD55C \uACFC\uBAA9 "}
          {extraCourses.length}
          {"\uAC1C \u00B7 \uD574\uB2F9 \uD559\uC0DD "}
          {withExtras.length}명
        </div>
        <button onClick={addExtraCourse} style={buttonStyle("primary")}>
          + 과목 추가
        </button>
      </div>
      {extraCourses.length === 0 && <EmptyPanel text="아직 입력한 과목이 없습니다. ‘+ 과목 추가’를 눌러 시작하세요." />}
      {extraCourses.map((x) => (
        <ExtraCourseCard
          key={x.id}
          x={x}
          results={results}
          allByKey={allByKey}
          groups={groups}
          subjectLookup={subjectLookup}
          updateExtraCourse={updateExtraCourse}
          removeExtraCourse={removeExtraCourse}
          addStudentToExtra={addStudentToExtra}
          removeStudentFromExtra={removeStudentFromExtra}
          onOpenCheck={onOpenCheck}
        />
      ))}
      {withExtras.length > 0 && (
        <div
          style={{
            ...CARD,
            marginTop: 16,
            overflow: "auto",
          }}
        >
          <div
            style={{
              fontSize: 13.5,
              fontWeight: 700,
              padding: "12px 14px 6px",
            }}
          >
            학생별 보기
          </div>
          <table
            style={{
              borderCollapse: "collapse",
              width: "100%",
              fontSize: 12.5,
              minWidth: 560,
            }}
          >
            <thead>
              <tr
                style={{
                  background: PAPER,
                }}
              >
                <Th>학번</Th>
                <Th>이름</Th>
                <Th>학급</Th>
                <Th>추가 이수 과목</Th>
                <Th>학점 합</Th>
              </tr>
            </thead>
            <tbody>
              {withExtras.map((r) => (
                <tr
                  key={r.key}
                  onClick={() => onOpenCheck(r.key)}
                  style={{
                    borderTop: `1px solid ${LINE}`,
                    cursor: "pointer",
                  }}
                >
                  <Td
                    style={{
                      color: MUTED,
                    }}
                  >
                    {r.studentId}
                  </Td>
                  <Td
                    style={{
                      fontWeight: 600,
                    }}
                  >
                    {r.name}
                  </Td>
                  <Td>{r.classLabel}</Td>
                  <Td
                    style={{
                      whiteSpace: "normal",
                    }}
                  >
                    {r.extraEntries.map((e) => `${e.name}(${e.credit})`).join(", ")}
                  </Td>
                  <Td
                    style={{
                      fontWeight: 700,
                    }}
                  >
                    {r.extraEntries.reduce((a, e) => a + e.credit, 0)}
                  </Td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default GraduationChecker;
