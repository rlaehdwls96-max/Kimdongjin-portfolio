# -*- coding: utf-8 -*-
r"""
index.html 에 "AI 활용 프로젝트 — REPEAT 발주 자동화" 항목을 삽입합니다.

사용법 (파워쉘):
    cd <Kimdongjin-portfolio 폴더>
    python apply_ai_section.py

동작:
  1) index.html 을 index_backup.html 로 백업
  2) 아래 12개 지점을 정확히 1회씩 치환
  3) 하나라도 못 찾으면 아무것도 바꾸지 않고 중단 (원본 안전)

삽입되는 것:
  - 히어로 상단 링크(경력기술서·GitHub 옆)에 "AI 활용 프로젝트" 추가
  - 푸터 버튼에도 동일 링크 추가
  - 상단 캐러셀(움직이는 카드)에 AI 카드 + 미니 인포그래픽 추가
  - 하단 All Work 그리드에 AI 카드 + 파이프라인 목업 추가
  - 라이트박스 상세: 파이프라인 인포그래픽 2종 + 실제 파이썬 코드/버그 수정 diff
"""
import io
import os
import shutil
import sys

TARGET = "index.html"
BACKUP = "index_backup.html"

# ---------------------------------------------------------------- 1) CSS
CSS_ANCHOR = "  /* ===== RESPONSIVE ===== */"
CSS_NEW = r"""  /* ===== AI 활용 프로젝트 — 코드 블록 / diff / 파이프라인 목업 ===== */
  .lb-code { border: 1px solid var(--card-border); border-radius: 12px; overflow: hidden; margin-top: 10px; background: #FBFAF8; }
  .lb-code-head { font-size: 10.5px; font-weight: 700; letter-spacing: 0.4px; padding: 8px 13px; border-bottom: 1px solid var(--card-border); color: #6d6459; background: var(--bg-soft); font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
  .lb-code pre { margin: 0; padding: 13px 14px; overflow-x: auto; font-size: 11.5px; line-height: 1.75; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; color: #2f2b26; white-space: pre; }
  .lb-code .c { color: #9a9288; }
  .lb-code .k { color: #8A421B; font-weight: 700; }
  .lb-code .s { color: #4A5B41; }
  .lb-code.bad { border-color: #EBD3C0; }
  .lb-code.bad .lb-code-head { background: #FBEFE9; border-bottom-color: #EBD3C0; color: #8A421B; }
  .lb-code.good { border-color: #CBD8C2; }
  .lb-code.good .lb-code-head { background: #EDF2E9; border-bottom-color: #CBD8C2; color: #4A5B41; }
  .lb-diff { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-top: 10px; }
  .mock-pipe { display: flex; align-items: stretch; gap: 4px; margin-top: 10px; }
  .mock-pipe-node { flex: 1; background: #ECEAF1; border-radius: 6px; padding: 9px 4px; text-align: center; display: flex; flex-direction: column; gap: 3px; justify-content: center; }
  .mock-pipe-node b { font-size: 8.5px; font-weight: 800; color: #5C5480; letter-spacing: 0.3px; }
  .mock-pipe-node span { font-size: 7.5px; color: #767676; line-height: 1.3; }
  .mock-pipe-node.hi { background: #F3EDE7; }
  .mock-pipe-node.hi b { color: #8A421B; }

"""

# ------------------------------------------------- 2) 모바일 대응(diff 1열)
RESP_ANCHOR = "    .lb-images { grid-template-columns: 1fr; }"
RESP_NEW = "    .lb-images { grid-template-columns: 1fr; }\n    .lb-diff { grid-template-columns: 1fr; }"

# ---------------------------------------------------------- 3) FIG 인포그래픽
FIG_ANCHOR = "const FIG = {\n\nsystem: `<svg"
FIG_NEW = r"""const FIG = {

aipipe: `<svg viewBox="0 0 920 310" xmlns="http://www.w3.org/2000/svg" font-family="inherit" role="img" aria-label="REPEAT 발주 자동화 파이프라인">
  <text x="22" y="24" font-size="13" font-weight="700" fill="#111214">REPEAT 발주 자동화 — 처리 파이프라인</text>
  <text x="898" y="24" font-size="11" fill="#767676" text-anchor="end">Python · openpyxl · SQLite · Streamlit</text>
  <g>
    <rect x="25" y="46" width="166" height="112" rx="12" fill="#FFFFFF" stroke="#E6E1DA"/>
    <text x="41" y="72" font-size="10" font-weight="800" letter-spacing="1" fill="#C4622D">01 · RAW</text>
    <text x="41" y="96" font-size="14.5" font-weight="800" fill="#111214">원본 계획서</text>
    <text x="41" y="119" font-size="10.5" fill="#6d6459">브라 · 팬티 · SL 3종</text>
    <text x="41" y="135" font-size="10.5" fill="#6d6459">병합셀 · 반복헤더 혼재</text>
    <path d="M196 100 l12 0 m-4 -4 l4 4 -4 4" stroke="#C9C2B8" stroke-width="1.6" fill="none" stroke-linecap="round"/>

    <rect x="201" y="46" width="166" height="112" rx="12" fill="#FFFFFF" stroke="#E6E1DA"/>
    <rect x="286" y="36" width="88" height="20" rx="10" fill="#C4622D"/>
    <text x="330" y="50" font-size="9.5" font-weight="800" fill="#FFFFFF" text-anchor="middle">오류 발견 지점</text>
    <text x="217" y="72" font-size="10" font-weight="800" letter-spacing="1" fill="#C4622D">02 · PARSE</text>
    <text x="217" y="96" font-size="14.5" font-weight="800" fill="#111214">정규화</text>
    <text x="217" y="119" font-size="10.5" fill="#6d6459">실제 월 컬럼만 선별</text>
    <text x="217" y="135" font-size="10.5" fill="#6d6459">절대월 타임라인 복원</text>
    <path d="M372 100 l12 0 m-4 -4 l4 4 -4 4" stroke="#C9C2B8" stroke-width="1.6" fill="none" stroke-linecap="round"/>

    <rect x="377" y="46" width="166" height="112" rx="12" fill="#FFFFFF" stroke="#E6E1DA"/>
    <text x="393" y="72" font-size="10" font-weight="800" letter-spacing="1" fill="#C4622D">03 · STORE</text>
    <text x="393" y="96" font-size="14.5" font-weight="800" fill="#111214">로컬 DB</text>
    <text x="393" y="119" font-size="10.5" fill="#6d6459">브랜드 × 시즌 ×</text>
    <text x="393" y="135" font-size="10.5" fill="#6d6459">품번 × 색상 단위 적재</text>
    <path d="M548 100 l12 0 m-4 -4 l4 4 -4 4" stroke="#C9C2B8" stroke-width="1.6" fill="none" stroke-linecap="round"/>

    <rect x="553" y="46" width="166" height="112" rx="12" fill="#FFFFFF" stroke="#E6E1DA"/>
    <text x="569" y="72" font-size="10" font-weight="800" letter-spacing="1" fill="#C4622D">04 · SIMULATE</text>
    <text x="569" y="96" font-size="14.5" font-weight="800" fill="#111214">소진 시뮬레이션</text>
    <text x="569" y="119" font-size="10.5" fill="#6d6459">현재고 − 월판매</text>
    <text x="569" y="135" font-size="10.5" fill="#6d6459">+ 확정 입고 · 12개월</text>
    <path d="M724 100 l12 0 m-4 -4 l4 4 -4 4" stroke="#C9C2B8" stroke-width="1.6" fill="none" stroke-linecap="round"/>

    <rect x="729" y="46" width="166" height="112" rx="12" fill="#FBF3EC" stroke="#EBD3C0"/>
    <text x="745" y="72" font-size="10" font-weight="800" letter-spacing="1" fill="#C4622D">05 · OUTPUT</text>
    <text x="745" y="96" font-size="14.5" font-weight="800" fill="#111214">발주 판단</text>
    <text x="745" y="119" font-size="10.5" fill="#6d6459">발주 필요월 · 수량</text>
    <text x="745" y="135" font-size="10.5" fill="#6d6459">대시보드 자동 표시</text>
  </g>

  <text x="22" y="200" font-size="12" font-weight="700" fill="#111214">브랜드별 리드타임 — 자재 세팅까지 역산한 최소 확보 기간</text>
  <g>
    <rect x="25" y="214" width="212" height="62" rx="12" fill="#F4F1EC" stroke="#E6E1DA"/>
    <text x="45" y="242" font-size="12.5" font-weight="800" fill="#111214">브라 (BR)</text>
    <text x="215" y="242" font-size="18" font-weight="800" fill="#8A421B" text-anchor="end">5개월</text>
    <text x="45" y="262" font-size="10" fill="#6d6459">세트 팬티도 동일 적용</text>

    <rect x="249" y="214" width="212" height="62" rx="12" fill="#F4F1EC" stroke="#E6E1DA"/>
    <text x="269" y="242" font-size="12.5" font-weight="800" fill="#111214">팬티 (PT)</text>
    <text x="439" y="242" font-size="18" font-weight="800" fill="#8A421B" text-anchor="end">4개월</text>
    <text x="269" y="262" font-size="10" fill="#6d6459">단품 · EC · 세일품 기준</text>

    <rect x="473" y="214" width="212" height="62" rx="12" fill="#F4F1EC" stroke="#E6E1DA"/>
    <text x="493" y="242" font-size="12.5" font-weight="800" fill="#111214">SL (FS·DW·HS)</text>
    <text x="663" y="242" font-size="18" font-weight="800" fill="#8A421B" text-anchor="end">4개월</text>
    <text x="493" y="262" font-size="10" fill="#6d6459">서브군 공통 기준</text>

    <rect x="697" y="214" width="198" height="62" rx="12" fill="#FBF3EC" stroke="#EBD3C0"/>
    <text x="717" y="242" font-size="12.5" font-weight="800" fill="#111214">안전 마진</text>
    <text x="873" y="242" font-size="18" font-weight="800" fill="#8A421B" text-anchor="end">+0.5</text>
    <text x="717" y="262" font-size="10" fill="#6d6459">리드타임 위 추가 확보</text>
  </g>
</svg>`,

aibug: `<svg viewBox="0 0 920 250" xmlns="http://www.w3.org/2000/svg" font-family="inherit" role="img" aria-label="원본 데이터 집계 오류 수정">
  <text x="22" y="24" font-size="13" font-weight="700" fill="#111214">발견한 오류 — 월별 판매 합계가 부풀려지고 있었다</text>
  <text x="22" y="44" font-size="10.5" fill="#767676">원본 엑셀은 12개 월 컬럼 뒤에 자체 소계·분기 컬럼을 이어 붙인다. 이를 월로 오인식해 중복 합산 중이었음</text>

  <text x="22" y="82" font-size="10.5" font-weight="800" letter-spacing="1" fill="#8A421B">수정 전 · 중복 합산</text>
  <rect x="22" y="92" width="560" height="34" rx="8" fill="#C4622D" fill-opacity="0.22" stroke="#C4622D" stroke-opacity="0.5"/>
  <text x="40" y="115" font-size="13" font-weight="800" fill="#8A421B">약 2.3배</text>
  <text x="596" y="115" font-size="10.5" fill="#6d6459">실제 수요보다 과대 계상 → 발주 판단 왜곡</text>

  <text x="22" y="156" font-size="10.5" font-weight="800" letter-spacing="1" fill="#4A5B41">수정 후 · 실제 12개월 합계</text>
  <rect x="22" y="166" width="243" height="34" rx="8" fill="#7A8F6E" fill-opacity="0.25" stroke="#7A8F6E" stroke-opacity="0.5"/>
  <text x="40" y="189" font-size="13" font-weight="800" fill="#4A5B41">기준값 (100%)</text>
  <text x="280" y="189" font-size="10.5" fill="#6d6459">원본 시트의 자체 합계 셀과 일치 확인</text>

  <text x="22" y="232" font-size="10.5" fill="#6d6459">혼입되던 컬럼 — 계(자체 소계) · 초도월 · 분기 생산량. 정규식으로 &quot;숫자+월&quot; 패턴만 인정하도록 수정</text>
</svg>`,

system: `<svg"""

# --------------------------------------------------- 4) MINI (캐러셀 카드용)
MINI_ANCHOR = "const MINI = {\ncase1: `<svg"
MINI_NEW = r"""const MINI = {
ai: `<svg viewBox="0 0 250 160" xmlns="http://www.w3.org/2000/svg" font-family="inherit" aria-hidden="true">
  <text x="20" y="34" font-size="9" font-weight="800" letter-spacing="1.2" fill="#C4622D">RAW → 발주 판단</text>
  <g>
    <rect x="20" y="46" width="38" height="40" rx="6" fill="#FFF" stroke="#DCD7CF"/><text x="39" y="70" font-size="8" font-weight="700" fill="#4a443c" text-anchor="middle">엑셀</text>
    <rect x="63" y="46" width="38" height="40" rx="6" fill="#FFF" stroke="#DCD7CF"/><text x="82" y="70" font-size="8" font-weight="700" fill="#4a443c" text-anchor="middle">파싱</text>
    <rect x="106" y="46" width="38" height="40" rx="6" fill="#FFF" stroke="#DCD7CF"/><text x="125" y="70" font-size="8" font-weight="700" fill="#4a443c" text-anchor="middle">DB</text>
    <rect x="149" y="46" width="38" height="40" rx="6" fill="#FFF" stroke="#DCD7CF"/><text x="168" y="67" font-size="8" font-weight="700" fill="#4a443c" text-anchor="middle">시뮬</text><text x="168" y="78" font-size="8" font-weight="700" fill="#4a443c" text-anchor="middle">레이션</text>
    <rect x="192" y="46" width="38" height="40" rx="6" fill="#C4622D" fill-opacity="0.18" stroke="#C4622D" stroke-opacity="0.45"/><text x="211" y="67" font-size="8" font-weight="800" fill="#8A421B" text-anchor="middle">발주</text><text x="211" y="78" font-size="8" font-weight="800" fill="#8A421B" text-anchor="middle">판단</text>
  </g>
  <text x="20" y="112" font-size="10.5" font-weight="800" fill="#4a443c">품번·색상 1,300여 건 자동 판별</text>
  <text x="20" y="130" font-size="8.5" fill="#767676">소진 예상월 · 발주 필요월 · 발주 수량</text>
  <text x="20" y="147" font-size="8.5" fill="#767676">Python · SQLite · Streamlit</text>
</svg>`,
case1: `<svg"""

# ------------------------------------------------------ 5) 캐러셀 카드 목록
CARD_ANCHOR = "  { key:'system', badge:'Working System', v:'v8', label:'SYSTEM', title:'일하는 구조', desc:'리피트·단종 판단 게이트 · 공급계획 · 아카이브', stat:'5단계 판단 게이트' },"
CARD_NEW = CARD_ANCHOR + "\n  { key:'ai', badge:'AI 활용', v:'v5', label:'AI', title:'AI 활용 — 발주 자동화', desc:'엑셀 원본 파싱부터 발주 시점 역산까지 직접 구현', stat:'품번·색상 1,300여 건 자동 판별' },"

# --------------------------------------------------------- 6) 케이스 본문
CASE_ANCHOR = "  /* ─────────── DATA SIDE PROJECT ─────────── */"
CASE_NEW = r"""  /* ─────────── AI 활용 — REPEAT 발주 자동화 ─────────── */
  ai: {
    eyebrow:'AI 활용 · Side Project',
    title:'매달 손으로 대조하던 발주 판단을 코드로 옮기다',
    period:'2026.08 – 진행 중 · 개인 프로젝트 · 기획 · 설계 · 구현 전담',
    images:[ { fig:'aipipe' }, { fig:'aibug' } ],
    colors:['#F0ECF4','#F3EDE7'],
    narrative:`
      <div class="lb-sec">
        <div class="lb-sec-label">Situation — 매달 반복되는 대조 작업</div>
        <div class="lb-body">브랜드별 REPEAT 계획서는 브라·팬티·SL 세 파일에, 시즌마다 새 시트로 쌓입니다. 품번과 색상별로 현재고와 월별 판매를 눈으로 대조해 "언제 얼마나 더 넣을지"를 판단해야 했고, <strong>같은 작업이 매달 반복</strong>됐습니다. 판단 자체보다 대조에 시간이 더 들어가는 구조였습니다.</div>
      </div>

      <div class="lb-sec">
        <div class="lb-sec-label">Action 01 — 원본을 열어보고 나서야 찾은 오류</div>
        <div class="lb-body">파싱 결과가 이상해 원본 엑셀을 셀 좌표 단위로 직접 열어봤습니다. 원인은 계산식이 아니라 <strong>원본의 열 구조</strong>였습니다. 12개 월 컬럼 뒤에 붙어 있는 계(자체 소계)·초도월·분기 생산량까지 "월"로 인식돼, 실제 합계 위에 그 합계 자신이 다시 더해지고 있었습니다.</div>
        <div class="lb-diff">
          <div class="lb-code bad">
            <div class="lb-code-head">수정 전 — 헤더를 그대로 수용</div>
            <pre><span class="c"># 헤더 텍스트를 전부 월 컬럼으로 간주</span>
month_cols = [
    str(c).strip() <span class="k">if</span> c <span class="k">else</span> f<span class="s">"col{i}"</span>
    <span class="k">for</span> i, c <span class="k">in</span> enumerate(row[COL_MONTH_START:])
]</pre>
          </div>
          <div class="lb-code good">
            <div class="lb-code-head">수정 후 — 정규식으로 월 컬럼만 선별</div>
            <pre>MONTH_HEADER_PATTERN = re.compile(r<span class="s">"^\\d{1,2}월$"</span>)

raw_cols = [str(c).strip() <span class="k">if</span> c <span class="k">else</span> <span class="s">""</span> <span class="k">for</span> c <span class="k">in</span> row[COL_MONTH_START:]]
month_cols = [c <span class="k">for</span> c <span class="k">in</span> raw_cols <span class="k">if</span> MONTH_HEADER_PATTERN.match(c)]</pre>
          </div>
        </div>
        <div class="lb-note">수정 후 값이 원본 시트의 자체 합계 셀과 일치하는지 품번 단위로 대조해 검증했습니다. 이 밖에도 품번 자리에 컵 범위 텍스트가 잘못 채워지는 오염 등 <b>원본 데이터 이슈 4건</b>을 찾아 정리했습니다.</div>
      </div>

      <div class="lb-sec">
        <div class="lb-sec-label">Action 02 — 재고 소진 시뮬레이션과 발주 시점 역산</div>
        <div class="lb-body">시트마다 시작월이 달라(4월 시작 · 7월 시작) 같은 "4월"이 서로 다른 해를 가리켰습니다. 이를 <strong>절대 캘린더월로 환산</strong>해 여러 시즌 스냅샷을 하나의 시계열로 병합한 뒤, 재고가 바닥나는 달을 찾고 리드타임만큼 역산해 발주 시점을 계산합니다.</div>
        <div class="lb-code">
          <div class="lb-code-head">재고 소진 시뮬레이션 · 발주 시점 역산</div>
          <pre><span class="c"># 월별로 재고를 굴려 처음 바닥나는 달을 찾는다</span>
<span class="k">for</span> m <span class="k">in</span> months:
    s = s - sales.get(m, 0) + prod.get(m, 0)   <span class="c"># 현재고 − 판매 + 확정 입고</span>
    <span class="k">if</span> s &lt;= 0 <span class="k">and</span> depletion <span class="k">is None</span>:
        depletion = m
        <span class="k">break</span>

lead = row.get(<span class="s">"리드타임_개월"</span>) <span class="k">or</span> 4.0     <span class="c"># 브라 5 / 팬티 4~5 / SL 4</span>
order_month = shift_month(depletion, -math.ceil(lead))

<span class="c"># 이미 지난 달로 계산되면 과거 날짜 대신 "지금 당장"으로 고정</span>
<span class="k">if</span> order_month &lt; today:
    order_month, state = today, <span class="s">"즉시발주 필요"</span>

<span class="c"># 발주 수량은 100장 단위 내림 (반올림 아님)</span>
order_qty = int(max(0, need - scheduled) // 100) * 100</pre>
        </div>
        <div class="lb-note">마지막 두 블록은 실제로 돌려보다가 나온 문제입니다. 발주 시점이 과거로 역산되는 논리 오류와, 발주 수량이 실무에서 쓰지 않는 단위로 떨어지는 문제를 각각 방어 코드로 정리했습니다.</div>
      </div>

      <div class="lb-sec">
        <div class="lb-sec-label">Result</div>
        <div class="lb-body">브랜드 3종 × 시즌 스냅샷을 통합해 <strong>품번·색상 1,300여 건</strong>의 소진 예상월·발주 필요월·발주 수량을 자동 산출합니다. 브랜드별 리드타임(브라 5개월 · 팬티 4~5개월 · SL 4개월)과 안전 마진까지 판단 규칙에 반영해, 매달 눈으로 대조하던 과정을 코드로 대체했습니다.</div>
      </div>

      <div class="lb-sec">
        <div class="lb-sec-label">AI를 쓰는 방식</div>
        <div class="lb-body">이 프로젝트는 Claude와 함께 만들었습니다. 제가 맡은 몫은 코드를 한 줄씩 쓰는 일이 아니라, <strong>실제 원본 데이터로 결과를 매번 검증하고, 틀린 가정을 짚어내고, 다음에 무엇을 고칠지 판단하는 일</strong>이었습니다. 위의 오류 4건은 모두 생성된 코드를 그대로 믿지 않고 원본과 대조하는 과정에서 나왔습니다. AI를 코드 생성기가 아니라 가설을 빠르게 검증하는 도구로 쓸 때 실무에서 무엇이 달라지는지 확인한 작업입니다.</div>
        <div class="lb-tags">
          <span class="lb-tag">Python · Pandas</span><span class="lb-tag">openpyxl</span><span class="lb-tag">SQLite</span><span class="lb-tag">Streamlit</span><span class="lb-tag">정규표현식 파싱</span><span class="lb-tag">시계열 정규화</span>
        </div>
        <div class="lb-note">비공개 저장소로 관리 중입니다. 실제 매출·재고 수치와 원본 데이터는 회사 보안 정책상 공개하지 않으며, 위 코드는 구조를 보여주기 위한 실제 구현부입니다.</div>
      </div>`,
    quote:'생성된 코드를 그대로 믿지 않고 원본과 대조한 덕분에, 오히려 원본 데이터의 오류를 먼저 찾았습니다.',
    metrics:[
      {num:'1,300+',desc:'자동 판별 품번·색상'},
      {num:'4건',desc:'발견·수정한 데이터 오류'},
      {num:'3종',desc:'브랜드 파일 통합'},
      {num:'12개월',desc:'소진 시뮬레이션 범위'}
    ]
  },

  /* ─────────── DATA SIDE PROJECT ─────────── */"""

# ------------------------------------------------------- 7) 히어로 상단 링크
HERO_ANCHOR = "    <a href=\"career.html\">경력기술서</a>"
HERO_NEW = ("    <a href=\"career.html\">경력기술서</a>\n"
            "    <span class=\"dot\"></span>\n"
            "    <a href=\"#\" onclick=\"event.preventDefault(); openCase('ai');\">AI 활용 프로젝트</a>")

# ------------------------------------------------------------ 8) 푸터 버튼
FOOT_ANCHOR = "        <a href=\"career.html\" class=\"f-pill\">경력기술서</a>"
FOOT_NEW = ("        <a href=\"career.html\" class=\"f-pill\">경력기술서</a>\n"
            "        <a href=\"#\" class=\"f-pill\" onclick=\"event.preventDefault(); expandFooterKeep(); openCase('ai');\">AI 활용 프로젝트</a>")

# ------------------------------------------------- 9) 하단 그리드에 카드 추가
GRID_ANCHOR = "  const restBlock = ['system','career','skills'].map(cardHtml).join('');"
GRID_NEW = "  const restBlock = ['system','ai','career','skills'].map(cardHtml).join('');"

# --------------------------------------------------------- 10) 목업 타입 매핑
MOCKMAP_ANCHOR = "  const mockTypes = { case1:'funnel', case2:'growth', case3:'share', system:'gates', career:'timeline', skills:'radar' };"
MOCKMAP_NEW = "  const mockTypes = { case1:'funnel', case2:'growth', case3:'share', system:'gates', ai:'pipeline', career:'timeline', skills:'radar' };"

# ------------------------------------------------------- 11) 목업 렌더 함수
MOCKFN_ANCHOR = "  if (type === 'radar') return `"
MOCKFN_NEW = r"""  if (type === 'pipeline') return `
    <div class="mock-preview"><div class="mock-dots"><span></span><span></span><span></span></div>
      <div class="mock-tag card-stat" style="position:static;align-self:flex-end;margin-bottom:2px;">1,300여 건 자동 판별</div>
      <div class="mock-pipe">
        <div class="mock-pipe-node"><b>RAW</b><span>원본 엑셀</span></div>
        <div class="mock-pipe-node"><b>PARSE</b><span>정규화</span></div>
        <div class="mock-pipe-node"><b>DB</b><span>SQLite</span></div>
        <div class="mock-pipe-node"><b>SIM</b><span>소진 예측</span></div>
        <div class="mock-pipe-node hi"><b>ORDER</b><span>발주 판단</span></div>
      </div></div>`;
  if (type === 'radar') return `"""

# --------------------------------------------------------- 12) 항목 수 문구
COUNT_ANCHOR = "    <span>8개 항목 · 케이스 3건 · 사이드 프로젝트 1건</span>"
COUNT_NEW = "    <span>9개 항목 · 케이스 3건 · 사이드 프로젝트 2건</span>"

# ------------------------------------------- 13) 푸터에서 케이스 열 때 헬퍼
HELPER_ANCHOR = "/* ============ CURSOR ============ */"
HELPER_NEW = r"""/* 푸터가 펼쳐진 상태에서 AI 케이스를 열면 라이트박스가 가려지지 않도록 접어둡니다 */
function expandFooterKeep() {
  const bar = document.getElementById('footerBar');
  if (bar && bar.classList.contains('expanded')) { bar.classList.remove('expanded'); bar.classList.add('collapsed'); }
}

/* ============ CURSOR ============ */"""

PATCHES = [
    ("CSS 블록 추가", CSS_ANCHOR, CSS_NEW + CSS_ANCHOR),
    ("모바일 diff 1열", RESP_ANCHOR, RESP_NEW),
    ("FIG 인포그래픽 2종", FIG_ANCHOR, FIG_NEW),
    ("MINI 캐러셀 비주얼", MINI_ANCHOR, MINI_NEW),
    ("캐러셀 카드 항목", CARD_ANCHOR, CARD_NEW),
    ("케이스 상세 본문", CASE_ANCHOR, CASE_NEW),
    ("히어로 상단 링크", HERO_ANCHOR, HERO_NEW),
    ("푸터 버튼", FOOT_ANCHOR, FOOT_NEW),
    ("하단 그리드 카드", GRID_ANCHOR, GRID_NEW),
    ("목업 타입 매핑", MOCKMAP_ANCHOR, MOCKMAP_NEW),
    ("목업 렌더 함수", MOCKFN_ANCHOR, MOCKFN_NEW),
    ("항목 수 문구", COUNT_ANCHOR, COUNT_NEW),
    ("푸터 헬퍼 함수", HELPER_ANCHOR, HELPER_NEW),
]


def main():
    if not os.path.exists(TARGET):
        print("[중단] %s 을 찾을 수 없습니다. 이 스크립트를 index.html 과 같은 폴더에 두고 실행하세요." % TARGET)
        return 1

    with io.open(TARGET, encoding="utf-8") as f:
        html = f.read()

    if "key:'ai'" in html:
        print("[중단] 이미 AI 항목이 들어있는 것 같습니다. 중복 적용을 막기 위해 종료합니다.")
        return 1

    # --- 먼저 전부 검사 (하나라도 실패하면 파일을 건드리지 않음) ---
    problems = []
    for name, anchor, _ in PATCHES:
        n = html.count(anchor)
        if n != 1:
            problems.append("  - %s : 기준 문구를 %d번 찾음 (1번이어야 함)" % (name, n))
    if problems:
        print("[중단] 아래 지점을 찾지 못했습니다. index.html 이 예상과 다릅니다.")
        print("\n".join(problems))
        print("\n원본은 전혀 변경되지 않았습니다.")
        return 1

    # --- 전부 확인됐으니 적용 ---
    for name, anchor, new in PATCHES:
        html = html.replace(anchor, new, 1)

    shutil.copyfile(TARGET, BACKUP)
    with io.open(TARGET, "w", encoding="utf-8") as f:
        f.write(html)

    print("[완료] %d개 지점을 수정했습니다." % len(PATCHES))
    print("       원본 백업 : %s" % BACKUP)
    print("       되돌리려면 : copy %s %s" % (BACKUP, TARGET))
    return 0


if __name__ == "__main__":
    sys.exit(main())
