// 덱편집기 화면. 서버에 묻는 길은 전부 POST(/api/...).
"use strict";

const 상태 = { 이름: null, 자료: null, 지금: 0, 탭: "장", 굽기형식: new Set(["pptx"]), 굽는중: false };
const $ = (s) => document.querySelector(s);

// ------------------------------------------------------------ 서버
async function 묻기(길, 몸 = {}) {
  const r = await fetch(`/api/${길}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(몸) });
  const j = await r.json();
  if (!r.ok) throw new Error(j.오류 || "알 수 없는 오류");
  return j;
}
function 알리기(말, 오류 = false) {
  const el = $("#알림");
  el.textContent = 말; el.className = "알림 보임" + (오류 ? " 오류" : "");
  clearTimeout(알리기.t); 알리기.t = setTimeout(() => (el.className = "알림"), 2600);
}

// ------------------------------------------------------------ 아이콘(직접 그린 선 그림)
const 아이콘 = {
  설계: '<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4M10 12h5M10 16h5"/>',
  장: '<rect x="3" y="6" width="14" height="10" rx="2"/><path d="M7 19h12a2 2 0 0 0 2-2V9"/>',
  재료: '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="m21 16-5-5-9 9"/>',
  테마: '<circle cx="12" cy="12" r="9"/><circle cx="8" cy="10" r="1.3"/><circle cx="12" cy="7.5" r="1.3"/><circle cx="16" cy="10" r="1.3"/><path d="M12 21a3 3 0 0 1 0-6h3"/>',
  꾸미기: '<path d="m12 3 1.8 4.6L18.5 9l-4.7 1.4L12 15l-1.8-4.6L5.5 9l4.7-1.4z"/><path d="M19 15l.8 2 2 .8-2 .8-.8 2-.8-2-2-.8 2-.8z"/>',
  감독관: '<path d="M12 3 4 6v6c0 4.5 3.4 8.3 8 9 4.6-.7 8-4.5 8-9V6z"/><path d="m8.5 12 2.5 2.5 4.5-5"/>',
  내보내기: '<path d="M12 15V3M7 8l5-5 5 5"/><path d="M5 13v6a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-6"/>',
};
const 탭들 = [
  ["설계", "덱"], ["장", "장"], ["재료", "재료"], ["테마", "테마"],
  ["꾸미기", "꾸미기"], ["감독관", "감독관"], ["내보내기", "내보내기"],
];

// 장 종류마다 작은 배치 그림(%) — 글자 목록 대신 그림으로 고른다
const 배치그림 = {
  표지: [["강", 6, 26, 6, 1.6], ["", 6, 32, 46, 9], ["", 6, 46, 34, 4], ["", 6, 82, 20, 3], ["그", 58, 11, 36, 78]],
  한마디: [["", 15, 30, 70, 10], ["", 25, 43, 50, 10], ["강", 46, 62, 8, 1.6], ["", 30, 70, 40, 4]],
  글그림: [["", 6, 9, 50, 7], ["강", 6, 20, 5, 1.2], ["", 6, 27, 34, 4], ["", 6, 35, 30, 4], ["", 6, 43, 32, 4], ["그", 48, 25, 46, 64]],
  글만: [["", 6, 9, 50, 7], ["강", 6, 20, 5, 1.2], ["", 6, 28, 80, 4], ["", 6, 37, 70, 4], ["", 6, 46, 76, 4], ["", 6, 55, 60, 4]],
  도표: [["", 6, 9, 50, 7], ["강", 6, 20, 5, 1.2], ["그", 10, 70, 8, 18], ["그", 22, 52, 8, 36], ["그", 34, 60, 8, 28], ["그", 46, 40, 8, 48], ["그", 58, 48, 8, 40], ["", 6, 88, 88, 0.6]],
  표: [["", 6, 9, 50, 7], ["강", 6, 20, 5, 1.2], ["강", 6, 28, 88, 7], ["", 6, 38, 88, 5], ["그", 6, 46, 88, 5], ["", 6, 54, 88, 5]],
  목차: [["", 6, 9, 40, 7], ["강", 6, 28, 6, 5], ["", 15, 29, 40, 4], ["강", 6, 42, 6, 5], ["", 15, 43, 34, 4], ["강", 6, 56, 6, 5], ["", 15, 57, 44, 4], ["강", 6, 70, 6, 5], ["", 15, 71, 30, 4]],
  비교: [["", 6, 9, 40, 7], ["그", 6, 25, 42, 64], ["그", 52, 25, 42, 64], ["", 10, 32, 16, 4], ["", 56, 32, 16, 4]],
  숫자: [["", 6, 9, 40, 7], ["그", 6, 30, 27, 52], ["그", 36.5, 30, 27, 52], ["그", 67, 30, 27, 52], ["강", 11, 42, 17, 10], ["강", 41.5, 42, 17, 10], ["강", 72, 42, 17, 10]],
  흐름: [["", 6, 9, 40, 7], ["그", 6, 35, 18, 37], ["강", 25.5, 50, 3, 4], ["그", 30, 35, 18, 37], ["강", 49.5, 50, 3, 4], ["그", 54, 35, 18, 37], ["강", 73.5, 50, 3, 4], ["그", 78, 35, 16, 37]],
  인용: [["강", 6, 12, 8, 14], ["", 12, 36, 76, 8], ["", 12, 48, 60, 8], ["", 16, 72, 30, 4]],
  사진가득: [["그", 0, 0, 100, 100], ["", 6, 76, 50, 8], ["", 6, 88, 34, 4]],
  결과해석: [["", 6, 9, 40, 7], ["그", 6, 25, 52, 64], ["강", 62, 25, 20, 8], ["강", 62, 41, 16, 8], ["", 62, 60, 32, 29]],
  사진격자: [["", 6, 9, 40, 7], ["그", 6, 25, 27, 30], ["그", 36.5, 25, 27, 30], ["그", 67, 25, 27, 30], ["그", 6, 60, 27, 30], ["그", 36.5, 60, 27, 30], ["그", 67, 60, 27, 30]],
  결론: [["", 6, 9, 40, 7], ["", 6, 28, 46, 5], ["", 6, 38, 40, 5], ["", 6, 48, 44, 5], ["그", 60, 25, 34, 64]],
  끝: [["강", 0, 0, 100, 100], ["", 30, 40, 40, 12], ["", 38, 62, 24, 4]],
};
function 배치그림html(종류) {
  const 칸 = 배치그림[종류] || 배치그림.글그림;
  return `<div class="배치그림">${칸.map(([k, x, y, w, h]) =>
    `<i class="${k}" style="left:${x}%;top:${y}%;width:${w}%;height:${h}%"></i>`).join("")}</div>`;
}

// ------------------------------------------------------------ 장 그림(1920×1080 을 줄여서)
function 장넣기(틀, html) {
  틀.innerHTML = html;
  맞추기(틀);
}
function 맞추기(틀) {
  const 장 = 틀.querySelector(":scope > .장");
  if (장) 장.style.transform = `scale(${틀.clientWidth / 1920})`;
}
const 크기감시 = new ResizeObserver((목록) => 목록.forEach((e) => 맞추기(e.target)));

// ------------------------------------------------------------ 그리기
function 그리기() {
  레일그리기(); 패널그리기(); 무대그리기(); 띠그리기(); 살피기그리기();
}

function 레일그리기() {
  const 감독 = 상태.자료?.감독 || [];
  const 멈 = 감독.filter((f) => f.수준 === "멈춤").length, 주 = 감독.filter((f) => f.수준 === "주의").length;
  $("#레일").innerHTML = 탭들.map(([키, 이름, 곧]) =>
    `<button data-탭="${키}" aria-selected="${상태.탭 === 키}" class="${곧 ? "곧" : ""}" title="${곧 ? "다음 단계에서 열려요" : 이름}">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round">${아이콘[키]}</svg>${이름}${
      키 === "감독관" && (멈 || 주) ? `<span class="레일표 ${멈 ? "멈춤" : "주의"}">${멈 || 주}</span>` : ""}</button>`).join("");
  $("#레일").querySelectorAll("button").forEach((b) => (b.onclick = () => { 상태.탭 = b.dataset.탭; 레일그리기(); 패널그리기(); }));
}

function 무대그리기() {
  const 장 = 상태.자료?.장들[상태.지금];
  const 틀 = $("#큰장");
  if (!장) { 틀.innerHTML = '<div class="빈칸">장이 없어요. 아래 + 로 더해 보세요.</div>'; return; }
  장넣기(틀, 장.html);
}

function 띠그리기() {
  const 띠 = $("#띠");
  const 장들 = 상태.자료?.장들 || [];
  띠.innerHTML = 장들.map((장, i) => `
    <div class="작은장 장틀" draggable="true" data-i="${i}" aria-current="${i === 상태.지금}" title="${i + 1}장 · ${장.종류}">
      <span class="번호">${i + 1}</span>${장.넘침 ? '<span class="넘침표">넘침</span>'
        : (상태.자료.감독 || []).some((f) => f.장 === i + 1 && f.수준 === "멈춤") ? '<span class="넘침표">멈춤</span>' : ""}${장.html}</div>`).join("")
    + '<button class="더하기" title="장 더하기">+</button>';
  띠.querySelectorAll(".작은장").forEach((el) => {
    맞추기(el); 크기감시.observe(el);
    el.onclick = () => { 상태.지금 = +el.dataset.i; 무대그리기(); 띠그리기(); 살피기그리기(); };
    el.ondragstart = (e) => { e.dataTransfer.setData("text/plain", el.dataset.i); el.classList.add("끌림"); };
    el.ondragend = () => el.classList.remove("끌림");
    el.ondragover = (e) => { e.preventDefault(); el.classList.add("놓을곳"); };
    el.ondragleave = () => el.classList.remove("놓을곳");
    el.ondrop = (e) => {
      e.preventDefault(); el.classList.remove("놓을곳");
      옮기기(+e.dataTransfer.getData("text/plain"), +el.dataset.i);
    };
  });
  띠.querySelector(".더하기").onclick = () => { 상태.탭 = "장"; 레일그리기(); 패널그리기(); };
}

function 옮기기(from, to) {
  if (from === to || isNaN(from)) return;
  const 장 = 상태.자료.설계.장;
  const [하나] = 장.splice(from, 1);
  장.splice(to, 0, 하나);
  상태.지금 = to;
  저장하기(true);
}

// ------------------------------------------------------------ 패널(탭 안은 전부 타일)
function 패널그리기() {
  const p = $("#패널");
  const 자료 = 상태.자료;
  if (!자료) { p.innerHTML = '<div class="빈칸">불러오는 중…</div>'; return; }
  const 그리개 = { 설계: 패널_설계, 장: 패널_장, 재료: 패널_재료, 테마: 패널_테마, 꾸미기: 패널_꾸미기, 감독관: 패널_감독관,
    내보내기: 패널_내보내기 }[상태.탭] || 패널_곧;
  그리개(p, 자료);
}

async function 패널_설계(p, 자료) {
  const 탭 = 상태.탭;
  const { 프로젝트 } = await 묻기("목록");
  if (상태.탭 !== 탭) return;           // 기다리는 사이 다른 탭을 눌렀으면 그 탭을 덮지 않는다
  const 휴지통 = 자료.설계.휴지통 || [];
  p.innerHTML = `<h2>새 덱</h2><p class="설명">회사 템플릿 PPTX(v1)를 넣으면 템플릿은 그대로 두고 장을 고치거나 더해서 v2로 저장해요.</p>
    <input id="새덱이름" placeholder="덱 이름 (예: A사 2차 미팅)" style="margin-bottom:8px">
    <div class="타일들">
      <button class="타일" id="빈덱"><span class="이름">빈 덱</span><span class="작은">처음부터 만들기</span></button>
      <label class="타일" id="v1덱"><input type="file" accept=".pptx" hidden id="v1파일"><span class="이름">PPTX v1에서 시작</span><span class="작은">템플릿 유지 → v2로 저장</span></label>
    </div>
    <div class="줄단추" style="margin:8px 0 18px"><input id="v1경로" placeholder="또는 v1 경로: D:\\보고\\중간보고_v1.pptx"><button class="작은단추" id="v1경로넣기">시작</button></div>
    <h2>덱</h2><p class="설명">열 덱을 골라요. 프로젝트 폴더에 있는 덱이 모두 보여요.</p>
    <div class="타일들 한줄">${프로젝트.map((x) => `
      <button class="타일" data-이름="${x.이름}" aria-pressed="${x.이름 === 상태.이름}">
        <span class="이름">${x.이름}</span><span class="작은">${x.장수}장 · ${x.고친때}</span></button>`).join("")}</div>
    <h2 style="margin-top:18px">휴지통</h2><p class="설명">뺀 장은 지우지 않고 여기에 남아요. 누르면 원래 자리로 되살려요.</p>
    ${휴지통.length ? `<div class="타일들 한줄">${휴지통.map((장, i) => `
      <button class="타일" data-되살리기="${i}"><span class="이름">${글(장.제목 || 장.큰말 || 장.인용 || (장.원본 ? `원본 ${장.원본}장` : 장.종류))}</span>
        <span class="작은">${글(장.종류)} · 되살리기</span></button>`).join("")}</div>` : '<div class="빈칸">비어 있어요</div>'}
    <details style="margin-top:18px"><summary class="작은">설계 원문 보기(고급)</summary>
      <p class="설명">화면에서 고친 내용도 여기에 저장돼요(이전 판은 .이전 폴더에 남아요).</p><pre class="설계글"></pre></details>`;
  p.querySelectorAll("[data-되살리기]").forEach((b) => (b.onclick = () => {
    const 설계 = 상태.자료.설계, 장 = 설계.휴지통.splice(+b.dataset.되살리기, 1)[0];
    const 자리 = Math.min(장._뺀자리 ?? 설계.장.length, 설계.장.length);
    delete 장._뺀자리;
    설계.장.splice(자리, 0, 장);
    상태.지금 = 자리; 저장하기(true); 패널그리기(); 알리기(`${자리 + 1}장으로 되살렸어요`);
  }));
  p.querySelector(".설계글").textContent = JSON.stringify(자료.설계, null, 2);
  p.querySelectorAll(".타일[data-이름]").forEach((b) => (b.onclick = () => 열기(b.dataset.이름)));
  const 이름칸 = p.querySelector("#새덱이름");
  const 새로 = async (몸) => {
    const 덱이름 = 이름칸.value.trim() || 몸.기본이름;
    if (!덱이름) return 알리기("덱 이름을 넣어 주세요", true);
    알리기("새 덱을 만드는 중…");
    try { const r = await 묻기("새덱", { 덱이름, ...몸, 기본이름: undefined }); await 열기(r.이름); 알리기("새 덱을 열었어요"); }
    catch (e) { 알리기(e.message, true); }
  };
  p.querySelector("#빈덱").onclick = () => 새로({});
  p.querySelector("#v1파일").onchange = async (e) => {
    const f = e.target.files[0]; if (!f) return;
    const 내용 = await new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result.split(",")[1]); r.readAsDataURL(f); });
    새로({ 파일이름: f.name, 내용, 기본이름: f.name.replace(/[_ -]?v\d+\.pptx$|\.pptx$/i, "") });
  };
  p.querySelector("#v1경로넣기").onclick = () => {
    const 경로 = p.querySelector("#v1경로").value.trim();
    if (경로) 새로({ 경로, 기본이름: 경로.split(/[\\/]/).pop().replace(/[_ -]?v\d+\.pptx$|\.pptx$/i, "") });
  };
}

// 견본 글: 모두 '넣으세요'가 들어 있어서, 안 바꾸고 구우면 감독관이 멈춘다
const 견본 = { 표지: { 큰말: "발표 제목을 넣으세요", 작은말: "부제를 넣으세요" },
  한마디: { 큰말: "한 문장으로 말할 결론을 넣으세요" },
  글그림: { 제목: "제목을 넣으세요", 글: ["요점을 넣으세요"] },
  글만: { 제목: "제목을 넣으세요", 글: ["요점을 넣으세요"] },
  도표: { 제목: "제목을 넣으세요", 모양: "막대" }, 표: { 제목: "제목을 넣으세요" },
  목차: { 제목: "오늘 이야기할 것", 항목: ["첫 이야기를 넣으세요", "둘째 이야기를 넣으세요", "셋째 이야기를 넣으세요"] },
  비교: { 제목: "제목을 넣으세요", 왼제목: "왼쪽 이름을 넣으세요", 왼글: ["요점을 넣으세요"], 오른제목: "오른쪽 이름을 넣으세요", 오른글: ["요점을 넣으세요"] },
  숫자: { 제목: "제목을 넣으세요", 숫자들: ["92% | 무엇의 값인지 넣으세요", "3배 | 무엇의 값인지 넣으세요"] },
  흐름: { 제목: "제목을 넣으세요", 단계: ["첫 단계를 넣으세요", "둘째 단계를 넣으세요", "셋째 단계를 넣으세요"] },
  인용: { 인용: "들은 말을 그대로 넣으세요", 누가: "누가, 언제 했는지 넣으세요" },
  사진가득: { 제목: "사진이 말하는 것을 넣으세요" },
  결과해석: { 제목: "제목을 넣으세요", 숫자들: ["값 | 설명을 넣으세요"], 해석: "이 결과가 뜻하는 것을 넣으세요" },
  사진격자: { 제목: "제목을 넣으세요", 그림들: [], 라벨들: [] },
  결론: { 제목: "정리", 결론: ["결론을 넣으세요"], 다음: ["다음에 할 일을 넣으세요"] },
  끝: { 큰말: "감사합니다" } };

function 패널_장(p, 자료) {
  const 바탕 = 자료.설계.바탕;
  const 쓴원본 = new Set(자료.설계.장.filter((x) => x.종류 === "원본장").map((x) => +x.원본));
  const v1쪽 = 바탕 ? (자료.재료.find((m) => m.아이디 === 바탕.재료)?.쪽보기 || []) : [];
  p.innerHTML = `<h2>장 더하기</h2><p class="설명">누르면 지금 장 뒤에 들어가요.${바탕 ? " 새 장은 v1 템플릿의 배경·로고를 입고 PPTX에 들어가요." : ""}</p>
    ${바탕 ? `<h2 style="font-size:13px;margin-top:6px">v1 장</h2><div class="타일들">${v1쪽.map((s, i) => `
      <button class="타일" data-원본="${i + 1}" ${쓴원본.has(i + 1) ? 'aria-pressed="true" disabled title="이미 덱에 있어요(한 장은 한 번만 넣을 수 있어요)"' : ""}><div class="장틀 그림칸"><img src="${s}" alt=""></div>
      <span class="작은">원본 ${i + 1}장${쓴원본.has(i + 1) ? " · 쓰는 중" : ""}</span></button>`).join("")}</div>
      <h2 style="font-size:13px;margin-top:14px">새 장</h2>` : ""}
    <div class="타일들">${Object.entries(자료.장종류).filter(([, v]) => !v.추가안함).map(([k, v]) => `
      <button class="타일" data-종류="${k}">${배치그림html(k)}<span class="이름">${k}</span><span class="작은">${v.설명}</span></button>`).join("")}</div>`;
  p.querySelectorAll("[data-원본]").forEach((b) => (b.onclick = () => {
    const 목록 = 상태.자료.설계.장, 자리 = Math.min(상태.지금 + 1, 목록.length);
    목록.splice(자리, 0, { 종류: "원본장", 원본: +b.dataset.원본 });
    상태.지금 = 자리; 저장하기(true); 알리기(`원본 ${b.dataset.원본}장을 넣었어요`);
  }));
  p.querySelectorAll(".타일[data-종류]").forEach((b) => (b.onclick = () => {
    const 새장 = { 종류: b.dataset.종류, ...structuredClone(견본[b.dataset.종류] || {}) };
    const 목록 = 상태.자료.설계.장;   // 저장 뒤 바뀐 자료를 써야 한다(옛 것을 붙잡지 않기)
    const 자리 = Math.min(상태.지금 + 1, 목록.length);
    목록.splice(자리, 0, 새장);
    상태.지금 = 자리;
    저장하기(true);
    알리기(`${b.dataset.종류} 장을 더했어요`);
  }));
}

// ---- 재료: PDF 회의록·PPTX v1·엑셀·그림을 넣고, 꺼낸 것을 장에 넣는다
function 패널_재료(p, 자료) {
  p.innerHTML = `<h2>재료</h2><p class="설명">파일을 여기에 끌어다 놓거나, 이 PC의 경로를 붙여 넣어요. 원본은 그대로 두고 그림·글·표로 풀어 둬요.</p>
    <label class="놓는곳" id="놓는곳"><input type="file" id="파일고르기" multiple hidden>
      <b>파일 끌어다 놓기</b><span class="작은">또는 눌러서 고르기 · PDF · PPTX · 엑셀/CSV · 그림 · mp4 · Origin · 글</span></label>
    <div class="줄단추" style="margin:10px 0 16px"><input id="경로칸" placeholder="예: D:\\회의록\\A사_미팅.pdf"><button class="작은단추" id="경로넣기">넣기</button></div>
    <div id="재료목록">${자료.재료.length ? "" : '<div class="빈칸">아직 넣은 재료가 없어요.</div>'}</div>`;
  const 칸 = p.querySelector("#재료목록");
  자료.재료.slice().reverse().forEach((m) => 칸.appendChild(재료카드(m)));
  const 놓는곳 = p.querySelector("#놓는곳");
  놓는곳.ondragover = (e) => { e.preventDefault(); 놓는곳.classList.add("위"); };
  놓는곳.ondragleave = () => 놓는곳.classList.remove("위");
  놓는곳.ondrop = (e) => { e.preventDefault(); 놓는곳.classList.remove("위"); 올리기([...e.dataTransfer.files]); };
  p.querySelector("#파일고르기").onchange = (e) => 올리기([...e.target.files]);
  p.querySelector("#경로넣기").onclick = async () => {
    const 경로 = p.querySelector("#경로칸").value.trim();
    if (!경로) return;
    await 재료로(() => 묻기("재료넣기", { 이름: 상태.이름, 경로 }), 경로.split(/[\\/]/).pop());
  };
}

async function 재료로(일, 이름) {
  알리기(`${이름} 푸는 중…`);
  try { 상태.자료 = await 일(); 패널그리기(); 살피기그리기(); 알리기(`${이름}을(를) 재료함에 넣었어요`); }
  catch (e) { 알리기(e.message, true); }
}

async function 올리기(파일들) {
  for (const f of 파일들) {
    const 내용 = await new Promise((ok, no) => {
      const r = new FileReader(); r.onload = () => ok(r.result.split(",")[1]); r.onerror = no; r.readAsDataURL(f);
    });
    await 재료로(() => 묻기("재료넣기", { 이름: 상태.이름, 파일이름: f.name, 내용 }), f.name);
  }
}

function 재료카드(m) {
  const el = document.createElement("div");
  el.className = "재료카드";
  el.innerHTML = `<div class="재료머리"><span class="꼬리표">${m.종류}</span><b title="${글(m.원래이름)}">${글(m.원래이름)}</b></div>
    <div class="작은">${글(m.꺼낸것 || "")}</div>
    ${(m.알림 || []).map((a) => `<div class="넘침알림">${글(a)}</div>`).join("")}
    ${m.쪽보기.length ? `<div class="소제목">쪽</div><div class="쪽띠">${m.쪽보기.map((s, i) => `<img src="${s}" alt="${i + 1}쪽" title="${i + 1}쪽">`).join("")}</div>` : ""}
    ${m.그림보기.length ? `<div class="소제목">그림 <span class="작은">눌러서 지금 장에 · 큰 장으로 끌어도 돼요</span></div>
      <div class="그림고르기">${m.그림보기.map((g) => `<button draggable="true" data-재료그림="${글(g.경로)}" title="${글(g.경로.split("/").pop())}"><img src="${g.보기}" alt=""></button>`).join("")}</div>` : ""}
    ${m.글.length ? `<div class="소제목">글 <span class="작은">골라서 장에 넣어요</span></div>
      <div class="글덩어리들">${m.글.map((g, i) => `<label class="글덩어리"><input type="checkbox" data-i="${i}"><span><em>${g.쪽}쪽</em> ${글(g.글)}</span></label>`).join("")}</div>
      <div class="줄단추"><button class="작은단추" data-할="지금장">지금 장 글에 더하기</button><button class="작은단추" data-할="새장">새 글만 장으로</button></div>` : ""}
    ${m.표.length ? `<div class="소제목">표</div>${m.표.map((t) => `<div class="표줄"><span><b>${글(t.이름)}</b> <span class="작은">${t.머리.length}열 · ${t.줄수}줄</span></span>
      <span><button class="작은단추" data-표="${글(t.경로)}" data-종류="도표">도표 장</button><button class="작은단추" data-표="${글(t.경로)}" data-종류="표">표 장</button></span></div>`).join("")}` : ""}`;
  el.querySelectorAll("[data-재료그림]").forEach((b) => {
    b.onclick = () => 그림넣기(b.dataset.재료그림);
    b.ondragstart = (e) => e.dataTransfer.setData("text/재료그림", b.dataset.재료그림);
  });
  el.querySelectorAll("[data-할]").forEach((b) => (b.onclick = () => {
    const 고른 = [...el.querySelectorAll(".글덩어리 input:checked")].map((c) => m.글[+c.dataset.i].글);
    if (!고른.length) return 알리기("넣을 글을 먼저 골라 주세요", true);
    const 목록 = 상태.자료.설계.장;
    const 장 = 목록[상태.지금];
    if (b.dataset.할 === "지금장" && 장 && Array.isArray(장.글 ?? [])) {
      if (!("글" in (상태.자료.장종류[장.종류]?.칸 || []).reduce((o, c) => ({ ...o, [c.이름]: 1 }), {})))
        return 알리기("이 장에는 글 칸이 없어요. '새 글만 장으로'를 써 주세요", true);
      장.글 = [...(장.글 || []), ...고른];
    } else {
      목록.splice(상태.지금 + 1, 0, { 종류: "글만", 제목: "제목을 넣으세요", 글: 고른 });
      상태.지금 += 1;
    }
    저장하기(true); 알리기("글을 넣었어요");
  }));
  el.querySelectorAll("[data-표]").forEach((b) => (b.onclick = () => {
    const 목록 = 상태.자료.설계.장;
    목록.splice(상태.지금 + 1, 0, { 종류: b.dataset.종류, 제목: "제목을 넣으세요", 표: b.dataset.표, ...(b.dataset.종류 === "도표" ? { 모양: "막대" } : {}) });
    상태.지금 += 1; 저장하기(true); 알리기(`${b.dataset.종류} 장을 만들었어요`);
  }));
  return el;
}

function 그림넣기(경로) {
  const 장 = 상태.자료.설계.장[상태.지금];
  const 칸 = 상태.자료.장종류[장?.종류]?.칸 || [];
  const 여럿 = 칸.find((c) => c.형태 === "그림여럿");
  if (여럿) {
    const 목록 = (장[여럿.이름] ||= []);
    if (목록.length >= (여럿.최대 || 6)) return 알리기(`그림은 ${여럿.최대 || 6}장까지예요`, true);
    목록.push(경로); 저장하기(true); 알리기(`그림을 넣었어요 (${목록.length}장)`); return;
  }
  const 자리 = 칸.find((c) => c.형태 === "그림" && !장[c.이름]) || 칸.find((c) => c.형태 === "그림");
  if (!자리) return 알리기("이 장에는 그림 자리가 없어요. 글그림·표지·사진 장에 넣어 주세요", true);
  장[자리.이름] = 경로; 저장하기(true); 알리기("그림을 넣었어요");
}

async function 패널_테마(p, 자료) {
  const 표본 = 자료.설계.장.find((장) => 장.종류 !== "원본장") || { 종류: "표지", 큰말: "제목" };
  const 지금것 = typeof 자료.설계.스타일 === "object" && 자료.설계.스타일 && !(자료.스타일 in 자료.내스타일들)
    ? { [자료.스타일]: { 설명: "이 덱에 들어 있는 스타일", 원본: 자료.설계.스타일 } } : {};
  p.innerHTML = `<h2>해부</h2><p class="설명">마음에 드는 PPTX를 넣으면 색·글자체·제목 장식을 읽어 이 덱에 바로 입혀요. 넣은 PPTX는 고치지 않아요.</p>
    <label class="놓는곳" id="해부곳"><input type="file" accept=".pptx" hidden id="해부파일">
      <b>참고 PPTX 놓기</b><span class="작은">또는 눌러서 고르기</span></label>
    <div class="줄단추" style="margin:8px 0 4px"><input id="해부경로" placeholder="또는 경로: D:\\자료\\좋은발표.pptx"><button class="작은단추" id="해부경로넣기">해부</button></div>
    <div id="해부결과"></div>
    <h2 style="margin-top:18px">테마</h2><p class="설명">지금 덱의 장 하나를 스타일마다 그려 봤어요. 누르면 덱 전체가 바뀌어요.</p>
    <div class="타일들" id="테마타일"></div>
    <h2 style="margin-top:18px">글자체</h2><p class="설명">PPTX에도 같은 글자체로 들어가요. 맑은 고딕은 회사 PC에 설치 없이 있어요.</p>
    <div class="타일들" id="글자체타일"></div>`;
  const 칸 = p.querySelector("#테마타일");
  const 타일 = (이름, 설명, 값, 글자체, 표) => {
    const b = document.createElement("button");
    b.className = "타일"; b.setAttribute("aria-pressed", 이름 === 자료.스타일);
    b.innerHTML = `<div class="장틀 그림칸"></div><span class="이름">${글(이름)}${표 ? ` <span class="꼬리표">${표}</span>` : ""}</span><span class="작은">${글(설명)}</span>`;
    b.onclick = () => {
      상태.자료.설계.스타일 = structuredClone(값);
      if (글자체) 상태.자료.설계.글자체 = 글자체;
      저장하기(true);
    };
    칸.appendChild(b);
    묻기("미리보기", { 이름: 상태.이름, 장: 표본, 스타일: 값, 글자체: 글자체 || 자료.글자체, 꾸밈: 자료.설계.꾸밈 }).then((r) => {
      const 틀 = b.querySelector(".장틀"); 장넣기(틀, r.html); 크기감시.observe(틀);
    });
  };
  for (const [이름, s] of Object.entries(지금것)) 타일(이름, s.설명, s.원본, null, "지금");
  for (const [이름, s] of Object.entries(자료.내스타일들 || {})) 타일(이름, s.설명, s.원본, s.원본.글자체, "해부");
  for (const [이름, s] of Object.entries(자료.스타일들)) 타일(이름, s.설명, 이름, null, "");
  p.querySelector("#글자체타일").innerHTML = 자료.글자체들.map((g) => `
    <button class="타일" data-글자체="${g}" aria-pressed="${g === 자료.글자체}">
      <span style="font:700 22px ${g === "Pretendard" ? "Pretendard" : "'Malgun Gothic','맑은 고딕'"},sans-serif">가나 Aa 1</span>
      <span class="이름">${g}</span></button>`).join("");
  p.querySelectorAll("[data-글자체]").forEach((b) => (b.onclick = () => { 상태.자료.설계.글자체 = b.dataset.글자체; 저장하기(true); }));
  const 해부 = async (몸) => {
    알리기("참고 PPTX를 해부하는 중…");
    try {
      await 저장마저();      // 아직 안 보낸 고침을 먼저 저장(해부가 설계를 다시 쓴다)
      const r = await 묻기("해부", { 이름: 상태.이름, ...몸 });
      상태.자료 = r; 그리기();
      $("#해부결과").innerHTML = `<div class="재료카드"><b>${글(r.해부.이름)} 스타일을 입혔어요</b>
        ${r.해부.읽은것.map((x) => `<span class="작은">· ${글(x)}</span>`).join("")}</div>`;
      알리기(`${r.해부.이름} 스타일을 입혔어요`);
    } catch (e) { 알리기(e.message, true); }
  };
  const 곳 = p.querySelector("#해부곳");
  const 파일로 = async (f) => {
    if (!f) return;
    if (!/\.pptx$/i.test(f.name)) return 알리기("해부는 PPTX 파일만 돼요", true);
    const 내용 = await new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result.split(",")[1]); r.readAsDataURL(f); });
    해부({ 파일이름: f.name, 내용 });
  };
  p.querySelector("#해부파일").onchange = (e) => 파일로(e.target.files[0]);
  곳.ondragover = (e) => { e.preventDefault(); 곳.classList.add("위"); };
  곳.ondragleave = () => 곳.classList.remove("위");
  곳.ondrop = (e) => { e.preventDefault(); 곳.classList.remove("위"); 파일로(e.dataTransfer.files[0]); };
  p.querySelector("#해부경로넣기").onclick = () => { const 경로 = p.querySelector("#해부경로").value.trim(); if (경로) 해부({ 경로 }); };
}

function 패널_내보내기(p) {
  const 형식 = [["pptx", "고칠 수 있는 PPTX", "글상자·도형·차트가 살아 있어요. 파워포인트에서 바로 고쳐요"],
    ["pdf", "PDF", "글자가 살아 있는 PDF"], ["png", "그림 묶음 (PNG zip)", "장마다 1920×1080 그림"],
    ["mp4", "발표 영상 (mp4)", "장마다 '시간'초(없으면 5초), 장 사이 부드럽게 넘김"],
    ["대본", "대본 (md)", "장마다 발표자 노트를 모은 글"], ["데이터", "데이터 (xlsx)", "도표·표에 쓴 표를 장마다 시트로"],
    ["소재", "그림 원본", "덱에 쓴 그림 원본을 화질 그대로 복사"]];
  const 바탕 = 상태.자료.설계.바탕;
  p.innerHTML = `<h2>내보내기</h2><p class="설명">고른 것을 한 번에 구워요. 결과는 프로젝트의 내보내기 폴더에 날짜별로 쌓여요(지우지 않아요).</p>
    ${바탕 ? `<div class="넘침알림" style="background:rgba(61,219,217,.08);border-color:rgba(61,219,217,.35);color:#bff3f2;margin-bottom:12px">
      바탕: <b>${글(바탕.원래이름)}</b> → 템플릿을 그대로 두고 v2를 만들어요. v1 파일은 바뀌지 않고, 바뀐 내용은 변경기록.md에 남아요.</div>` : ""}
    <div class="타일들 한줄">${형식.map(([k, 이름, 설명]) => `
      <button class="타일" data-형식="${k}" aria-pressed="${상태.굽기형식.has(k)}"><span class="이름">${이름}</span><span class="작은">${설명}</span></button>`).join("")}</div>
    <div class="줄단추" style="margin-top:14px"><button class="큰단추" id="굽기">굽기</button>
      <button class="작은단추" id="전부고르기">전부 고르기</button></div>
    <div class="결과" id="결과"></div>`;
  결과그리기();
  p.querySelectorAll("[data-형식]").forEach((b) => (b.onclick = () => {
    const k = b.dataset.형식;
    상태.굽기형식.has(k) ? 상태.굽기형식.delete(k) : 상태.굽기형식.add(k);
    b.setAttribute("aria-pressed", 상태.굽기형식.has(k));
  }));
  p.querySelector("#굽기").onclick = () => 굽기([...상태.굽기형식]);
  p.querySelector("#전부고르기").onclick = () => {
    형식.forEach(([k]) => 상태.굽기형식.add(k));
    p.querySelectorAll("[data-형식]").forEach((b) => b.setAttribute("aria-pressed", "true"));
  };
}

function 패널_꾸미기(p, 자료) {
  const 지금장 = 자료.설계.장[상태.지금];
  const 표본 = 지금장 && 지금장.종류 !== "원본장" ? 지금장
    : 자료.설계.장.find((장) => 장.그림 && 장.종류 !== "원본장") || 자료.설계.장.find((장) => 장.종류 !== "원본장");
  const 지금 = 자료.설계.꾸밈 || {};
  p.innerHTML = `<h2>꾸미기</h2><p class="설명">기본은 꾸밈 없음이에요. 고르면 덱 전체에 입혀져요. 사진 효과는 보이는 것만 바꾸고 원본 그림은 그대로 들어가요.</p>
    <div class="줄단추" style="margin-bottom:10px"><button class="작은단추" id="꾸밈없애기">모두 없음으로</button></div>
    ${표본 ? "" : '<div class="빈칸">꾸밀 장이 없어요.</div>'}
    ${Object.entries(자료.꾸밈들).map(([무리, 들]) => `<h2 style="margin-top:14px">${무리}</h2>
      <div class="타일들" data-무리="${무리}">${Object.entries(들).map(([값, 설명], i) => `
        <button class="타일" data-값="${값}" aria-pressed="${(지금[무리] || Object.keys(들)[0]) === 값}">
          <div class="장틀 그림칸"></div><span class="이름">${값}</span><span class="작은">${설명}</span></button>`).join("")}</div>`).join("")}`;
  const 바꾸기 = (무리, 값) => {
    const 꾸 = { ...(상태.자료.설계.꾸밈 || {}), [무리]: 값 };
    for (const [k, 들] of Object.entries(상태.자료.꾸밈들)) if (꾸[k] === Object.keys(들)[0]) delete 꾸[k];
    if (Object.keys(꾸).length) 상태.자료.설계.꾸밈 = 꾸; else delete 상태.자료.설계.꾸밈;
    저장하기(true);
  };
  p.querySelector("#꾸밈없애기").onclick = () => { delete 상태.자료.설계.꾸밈; 저장하기(true); };
  p.querySelectorAll("[data-무리]").forEach((칸) => 칸.querySelectorAll(".타일").forEach((b) => {
    b.onclick = () => 바꾸기(칸.dataset.무리, b.dataset.값);
    if (!표본) return;
    묻기("미리보기", { 이름: 상태.이름, 장: 표본, 스타일: 자료.설계.스타일 ?? 자료.스타일, 글자체: 자료.글자체,
      꾸밈: { ...지금, [칸.dataset.무리]: b.dataset.값 } }).then((r) => {
      const 틀 = b.querySelector(".장틀"); 장넣기(틀, r.html); 크기감시.observe(틀);
    });
  }));
}

const 고침이름 = { 담기로: "자르지 않고 전체 보이게", 장나누기: "두 장으로 나누기", 글만으로: "글만 장으로 바꾸기", 축0켜기: "0에서 시작하게", 장빼기: "이 장 빼기" };
function 패널_감독관(p, 자료) {
  const 감독 = 자료.감독 || [];
  const 무리 = [["멈춤", "멈춤", "이대로 내보내면 안 돼요. 굽기 전에 다시 물어요"], ["주의", "주의", "고치면 더 좋아요"], ["알림", "알림", "알아 두면 좋아요"]];
  p.innerHTML = `<h2>감독관</h2><p class="설명">고칠 때마다 다시 살펴요. 고침 단추는 눌러야만 바뀌고, 뺀 장은 휴지통에 남아요.</p>
    ${감독.length ? 무리.map(([k, 이름, 설명]) => { const 들 = 감독.filter((f) => f.수준 === k); return 들.length ? `
      <div class="소제목">${이름} ${들.length} <span class="작은" style="font-weight:400">· ${설명}</span></div>
      <div class="감독목록">${들.map((f) => `<div class="감독줄 ${k}">
        <div>${f.장 ? `<button class="작은단추" data-가기="${f.장}">${f.장}장</button>` : '<span class="꼬리표">덱</span>'} ${글(f.말)}</div>
        <div class="줄단추">${f.고침 ? `<button class="작은단추 고르기" data-고침="${f.고침}" data-규칙="${f.규칙}" data-장="${f.장}">${고침이름[f.고침] || f.고침}</button>` : ""}
          <button class="작은단추" data-무시="${f.규칙}" data-장="${f.장}" data-말="${글(f.말)}" title="다음부터 이 줄은 안 보여요">괜찮아요</button></div></div>`).join("")}</div>` : ""; }).join("")
      : '<div class="빈칸">찾은 것이 없어요. 구워도 좋아요 ✓</div>'}`;
  p.querySelectorAll("[data-가기]").forEach((b) => (b.onclick = () => { 상태.지금 = +b.dataset.가기 - 1; 무대그리기(); 띠그리기(); 살피기그리기(); }));
  const 보내기 = async (길, 몸) => {
    try {
      await 저장마저();
      const r = await 묻기(길, { 이름: 상태.이름, ...몸 });
      상태.자료 = r; 그리기(); if (r.말) 알리기(r.말);
    } catch (e) { 알리기(e.message, true); }
  };
  p.querySelectorAll("[data-고침]").forEach((b) => (b.onclick = () => 보내기("감독고치기", { 규칙: b.dataset.규칙, 장: +b.dataset.장, 고침: b.dataset.고침 })));
  p.querySelectorAll("[data-무시]").forEach((b) => (b.onclick = () => 보내기("감독무시", { 규칙: b.dataset.무시, 장: +b.dataset.장, 말: b.dataset.말 })));
}

function 패널_곧(p) {
  const 말 = { 재료: "PDF 회의록·PPTX v1·엑셀·Origin·사진을 넣고 끌어서 장에 넣는 곳이에요.",
    꾸미기: "배경·사진 필터·테두리·그림자를 타일로 골라 입히는 곳이에요.",
  }[상태.탭];
  p.innerHTML = `<h2>${상태.탭}</h2><p class="설명">${말}</p><div class="빈칸">다음 단계에서 열려요.</div>`;
}

// ------------------------------------------------------------ 살피는 칸
function 살피기그리기() {
  const s = $("#살피기");
  const 자료 = 상태.자료;
  const 장 = 자료?.설계.장[상태.지금];
  if (!장) { s.innerHTML = '<div class="빈칸">장을 고르면 여기서 고쳐요.</div>'; return; }
  const 틀 = 자료.장종류[장.종류] || { 칸: [] };
  const 넘침 = 자료.장들[상태.지금]?.넘침;
  s.innerHTML = `<h3>${상태.지금 + 1}장 <span class="꼬리표">${장.종류}</span></h3>
    ${넘침 ? '<div class="넘침알림">글이 상자보다 길어요. 가장 작은 글자까지 줄여도 넘쳐서 빨간 점선으로 표시했어요. 글을 줄이거나 장을 나눠 주세요.</div>' : ""}
    ${틀.칸.map((칸) => 칸입력(칸, 장)).join("")}
    ${장.종류 === "원본장" ? "" : `<div><div class="칸이름"><span>발표자 노트</span></div><textarea data-칸="노트" rows="3">${글(장.노트)}</textarea></div>`}
    <div class="줄단추"><button class="작은단추" data-할일="앞">← 앞으로</button><button class="작은단추" data-할일="뒤">뒤로 →</button>
      <button class="작은단추" data-할일="빼기" title="휴지통으로 옮겨요(지우지 않아요)">이 장 빼기</button></div>`;
  s.querySelectorAll("[data-칸]").forEach((el) => (el.oninput = () => {
    const 칸 = el.dataset.칸, 장 = 상태.자료.설계.장[상태.지금];
    장[칸] = el.dataset.형태 === "여러줄" ? el.value.split("\n").filter((x) => x.trim()) : el.value;
    한도표시(el);
    저장하기();
  }));
  s.querySelectorAll("[data-원본글]").forEach((el) => (el.oninput = () => {
    const 장 = 상태.자료.설계.장[상태.지금];
    const 원래 = 상태.자료.원본글[(+장.원본) - 1].글상자.find((x) => String(x.아이디) === el.dataset.원본글)?.글;
    장.고침 ||= {};
    if (el.value === 원래) delete 장.고침[el.dataset.원본글]; else 장.고침[el.dataset.원본글] = el.value;
    저장하기();
  }));
  s.querySelectorAll("select[data-칸]").forEach((el) => (el.onchange = () => {
    const 장 = 상태.자료.설계.장[상태.지금];
    장[el.dataset.칸] = el.value; 저장하기(true);
  }));
  s.querySelectorAll("[data-고르기]").forEach((b) => (b.onclick = () => {
    const 장 = 상태.자료.설계.장[상태.지금];
    장[b.dataset.고르기] = b.dataset.값; 저장하기(true);
  }));
  s.querySelectorAll("[data-그림]").forEach((b) => (b.onclick = () => {
    const 장 = 상태.자료.설계.장[상태.지금], 칸 = b.dataset.그림칸;
    if (b.dataset.여럿) {          // 여러 장: 누른 차례대로 넣고, 다시 누르면 뺀다
      const 목록 = (장[칸] ||= []), i = 목록.indexOf(b.dataset.그림);
      if (i >= 0) 목록.splice(i, 1); else if (목록.length < +b.dataset.여럿) 목록.push(b.dataset.그림);
      else return 알리기(`그림은 ${b.dataset.여럿}장까지예요`, true);
    } else if (b.dataset.그림) 장[칸] = b.dataset.그림; else delete 장[칸];
    저장하기(true); 살피기그리기();
  }));
  s.querySelectorAll("[data-탭가기]").forEach((b) => (b.onclick = () => { 상태.탭 = b.dataset.탭가기; 레일그리기(); 패널그리기(); }));
  s.querySelectorAll("[data-할일]").forEach((b) => (b.onclick = () => {
    const i = 상태.지금, 목록 = 상태.자료.설계.장;
    if (b.dataset.할일 === "앞" && i > 0) 옮기기(i, i - 1);
    if (b.dataset.할일 === "뒤" && i < 목록.length - 1) 옮기기(i, i + 1);
    if (b.dataset.할일 === "빼기") {
      (상태.자료.설계.휴지통 ||= []).push({ ...목록.splice(i, 1)[0], _뺀자리: i });
      상태.지금 = Math.max(0, i - 1); 저장하기(true); 알리기("휴지통으로 옮겼어요. 덱 탭의 휴지통에서 되살릴 수 있어요");
    }
  }));
}
const 글 = (v) => (v ?? "").toString().replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

function 재료안내(칸이름, 무엇) {
  return `<div><div class="칸이름"><span>${칸이름}</span></div>
    <div class="빈칸" style="padding:12px">아직 ${무엇}이 없어요. 재료 탭에서 먼저 넣어 주세요.<br>
    <button class="작은단추" data-탭가기="재료" style="margin-top:8px">재료 탭으로</button></div></div>`;
}
function 칸입력(칸, 장) {
  if (칸.형태 === "원본글") {
    const 상자들 = 상태.자료.원본글[(+장.원본 || 1) - 1]?.글상자 || [];
    const 고침 = 장.고침 || {};
    return `<div class="넘침알림" style="background:rgba(61,219,217,.08);border-color:rgba(61,219,217,.35);color:#bff3f2">
      v1 장을 그대로 써요. 여기서 글을 고치면 v2 PPTX에 반영돼요(글자 모양은 v1 그대로). 가운데 그림은 v1을 구운 것이라 고친 글은 굽고 나서 보여요.</div>
      ${상자들.map((s) => `<div><div class="칸이름"><span>${글(s.이름)}</span><span>${고침[s.아이디] !== undefined ? "고침" : ""}</span></div>
        <textarea data-원본글="${s.아이디}" rows="${Math.min(6, s.글.split("\n").length + 1)}">${글(고침[s.아이디] ?? s.글)}</textarea></div>`).join("")
        || '<div class="빈칸">고칠 글상자가 없는 장이에요.</div>'}`;
  }
  if (칸.형태 === "표") {
    const 표들 = 상태.자료.재료.flatMap((m) => m.표.map((t) => ({ ...t, 재료: m.원래이름 })));
    if (!표들.length) return 재료안내(칸.이름, "엑셀·CSV");
    return `<div><div class="칸이름"><span>${칸.이름}</span><span>재료함의 표</span></div>
      <select data-칸="${칸.이름}"><option value="">고르기…</option>${표들.map((t) =>
        `<option value="${글(t.경로)}" ${장[칸.이름] === t.경로 ? "selected" : ""}>${글(t.재료)} · ${글(t.이름)}</option>`).join("")}</select></div>`;
  }
  if (칸.형태 === "고르기") {
    return `<div><div class="칸이름"><span>${칸.이름}</span></div><div class="줄단추">${칸.고를것.map((v) =>
      `<button class="작은단추 고르기" data-고르기="${칸.이름}" data-값="${v}" aria-pressed="${(장[칸.이름] || 칸.고를것[0]) === v}">${v}</button>`).join("")}</div></div>`;
  }
  if ((칸.형태 === "그림" || 칸.형태 === "그림여럿") && !상태.자료.소재.length) return 재료안내(칸.이름, "사진·그림");
  if (칸.형태 === "그림") {
    const 소재 = 상태.자료.소재, 지금 = 장[칸.이름];
    return `<div><div class="칸이름"><span>${칸.이름}${칸.필수 ? " *" : ""}</span><span>${소재.length}개 중 고르기</span></div>
      <div class="그림고르기"><button class="없음" data-그림="" data-그림칸="${칸.이름}" aria-pressed="${!지금}">없음</button>
      ${소재.map((m) => `<button data-그림="${글(m.경로)}" data-그림칸="${칸.이름}" aria-pressed="${지금 === m.경로}" title="${글(m.이름)} · ${m.너비}×${m.높이}"><img src="${m.보기}" alt=""></button>`).join("")}</div></div>`;
  }
  if (칸.형태 === "그림여럿") {
    const 소재 = 상태.자료.소재, 목록 = 장[칸.이름] || [], 최대 = 칸.최대 || 6;
    return `<div><div class="칸이름"><span>${칸.이름}${칸.필수 ? " *" : ""} (누른 차례대로)</span><span>${목록.length}/${최대}</span></div>
      <div class="그림고르기">${소재.map((m) => { const i = 목록.indexOf(m.경로);
        return `<button data-그림="${글(m.경로)}" data-그림칸="${칸.이름}" data-여럿="${최대}" aria-pressed="${i >= 0}" title="${글(m.이름)}" style="position:relative"><img src="${m.보기}" alt="">${i >= 0 ? `<span class="차례">${i + 1}</span>` : ""}</button>`; }).join("")}</div></div>`;
  }
  const 값 = 칸.형태 === "여러줄" ? (장[칸.이름] || []).join("\n") : (장[칸.이름] || "");
  const 입력 = 칸.형태 === "여러줄"
    ? `<textarea data-칸="${칸.이름}" data-형태="여러줄" data-한도="${칸.한도 || ""}" rows="${(칸.줄수 || 3) + 1}">${글(값)}</textarea>`
    : `<input data-칸="${칸.이름}" data-한도="${칸.한도 || ""}" value="${글(값)}">`;
  return `<div><div class="칸이름"><span>${칸.이름}${칸.필수 ? " *" : ""}${칸.형태 === "여러줄" ? " (한 줄에 하나)" : ""}</span><span class="세기"></span></div>${입력}</div>`;
}
function 한도표시(el) {
  const 한도 = +el.dataset.한도; const 세기 = el.parentElement.querySelector(".세기");
  if (!한도 || !세기) return;
  const 줄들 = el.value.split("\n"); const 가장긴 = Math.max(...줄들.map((x) => x.length));
  세기.textContent = `${가장긴}/${한도}자`; 세기.className = "세기" + (가장긴 > 한도 ? " 넘음" : "");
}

// ------------------------------------------------------------ 저장·굽기·열기
// 저장은 한 줄로 세워 차례로 보낸다. 응답이 오는 사이에 친 글을 잃지 않도록,
// 돌아온 화면 자료는 받되 설계는 지금 화면의 것(가장 새것)을 그대로 둔다.
let 저장대기 = null, 저장줄 = Promise.resolve(), 저장바로 = false;
function 저장하기(바로 = false) {
  $("#상태").textContent = "저장 중…";
  저장바로 = 저장바로 || 바로;
  clearTimeout(저장대기);
  저장대기 = setTimeout(저장보내기, 바로 ? 0 : 350);
}
function 저장보내기() {
  저장대기 = null;
  const 바로 = 저장바로, 이름 = 상태.이름;
  저장바로 = false;
  저장줄 = 저장줄.then(async () => {
    try {
      const 자료 = await 묻기("저장", { 이름, 설계: 상태.자료.설계 });
      if (상태.이름 !== 이름) return;            // 그사이 다른 덱을 열었으면 그대로 둔다
      상태.자료 = { ...자료, 설계: 상태.자료.설계 };
      상태.지금 = Math.min(상태.지금, Math.max(0, 자료.장들.length - 1));
      레일그리기(); 무대그리기(); 띠그리기();
      if (바로) { 살피기그리기(); if (["테마", "꾸미기", "감독관"].includes(상태.탭)) 패널그리기(); }
      $("#상태").textContent = 저장대기 ? "저장 중…" : "저장됨";
    } catch (e) { 알리기(e.message, true); $("#상태").textContent = "저장 못 함"; }
  });
  return 저장줄;
}
// 굽기·고치기·해부·다른 덱 열기 전에: 기다리는 저장을 바로 보내고 끝날 때까지 기다린다
async function 저장마저() {
  if (저장대기) { clearTimeout(저장대기); 저장보내기(); }
  await 저장줄;
}
function 결과그리기() {
  const 결과칸 = $("#결과"), r = 상태.굽기결과;
  if (!결과칸) return;
  if (!r) { 결과칸.innerHTML = ""; return; }
  if (r.굽는중) { 결과칸.innerHTML = '<div class="빈칸">굽는 중… (장이 많으면 몇십 초 걸려요)</div>'; return; }
  const 점검 = (r.점검 || []).map((f) => `<div class="넘침알림">완성 점검: ${글(f.말)}</div>`).join("");
  const 목록 = r.파일.map((f) => `<a href="${f.주소}" target="_blank" download>${글(f.이름)}<span>받기 ↓</span></a>`).join("");
  결과칸.innerHTML = `<div class="작은" title="${글(r.전체폴더 || "")}">저장한 곳: ${글(r.전체폴더 || r.폴더)}</div>
    <div class="줄단추"><button class="작은단추" id="폴더열기">폴더 열기</button></div>${목록}
    ${r.알림.map((m) => `<div class="넘침알림">${글(m)}</div>`).join("")}${점검 || '<div class="작은">완성 점검: 장 수·글상자 이상 없음 ✓</div>'}`;
  결과칸.scrollIntoView({ block: "nearest", behavior: "smooth" });     // 작은 화면에서도 결과가 보이게
  결과칸.querySelector("#폴더열기").onclick = () => 묻기("폴더열기", { 이름: 상태.이름, 폴더: r.폴더 }).catch((e) => 알리기(e.message, true));
}
async function 굽기(형식, 바로받기 = false) {
  if (!형식.length) return 알리기("구울 것을 하나 이상 골라 주세요", true);
  if (상태.굽는중) return;
  상태.굽는중 = true; $("#빨리굽기").disabled = true;
  const 덱 = 상태.이름;
  상태.굽기결과 = { 굽는중: true }; 결과그리기();
  try {
    await 저장마저();          // 방금 고친 글까지 굽는다
    let r = await 묻기("굽기", { 이름: 덱, 형식 });
    if (r.멈춤) {
      const 말 = r.멈춤.map((f) => `· ${f.장 ? f.장 + "장: " : ""}${f.말}`).join("\n");
      if (!confirm(`감독관이 멈췄어요.\n\n${말}\n\n그래도 구울까요? (취소하면 감독관 탭에서 고칠 수 있어요)`)) {
        상태.굽기결과 = null; 결과그리기(); 상태.탭 = "감독관"; 레일그리기(); 패널그리기(); return;
      }
      r = await 묻기("굽기", { 이름: 덱, 형식, 강행: true });
    }
    상태.굽기결과 = 상태.이름 === 덱 ? r : null;   // 탭을 바꿔도 남아 있다가 내보내기 탭에서 다시 보인다
    결과그리기();
    if (바로받기 && r.파일[0]) location.href = r.파일[0].주소;     // 위쪽 'PPTX 굽기'는 바로 내려받기
    알리기(`구웠어요: ${r.파일.length}개 파일 · 내보내기 탭에서 받거나 폴더를 열 수 있어요`);
  } catch (e) { 알리기(e.message, true); 상태.굽기결과 = null; 결과그리기(); }
  finally { 상태.굽는중 = false; $("#빨리굽기").disabled = false; }
}

async function 열기(이름) {
  if (상태.이름) await 저장마저();
  상태.이름 = 이름; 상태.지금 = 0; 상태.굽기결과 = null;
  상태.자료 = await 묻기("열기", { 이름 });
  history.replaceState(null, "", "?덱=" + encodeURIComponent(이름));   // 새로 고쳐도 이 덱이 열리게
  $("#덱이름").textContent = 상태.자료.설계.이름 || 이름;
  $("#상태").textContent = "열었어요";
  document.title = `${상태.자료.설계.이름 || 이름} · 덱편집기`;
  그리기();
}

async function 시작() {
  크기감시.observe($("#큰장"));
  const 무대 = $("#무대");
  무대.ondragover = (e) => { if ([...e.dataTransfer.types].includes("text/재료그림")) e.preventDefault(); };
  무대.ondrop = (e) => { const g = e.dataTransfer.getData("text/재료그림"); if (g) { e.preventDefault(); 그림넣기(g); } };
  $("#빨리굽기").onclick = () => 굽기(["pptx"], true);
  레일그리기();
  try {
    const { 프로젝트 } = await 묻기("목록");
    if (!프로젝트.length) { $("#패널").innerHTML = '<div class="빈칸">프로젝트 폴더에 덱설계.json 이 없어요.</div>'; return; }
    await 열기(new URLSearchParams(location.search).get("덱") || 프로젝트[0].이름);
  } catch (e) { 알리기(e.message, true); }
}
시작();
