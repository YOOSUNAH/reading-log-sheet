/**
 * [2026 나의 독서 기록장 & 서재 v4 - 심플 5-Star 인터랙션 & 그린 테마]
 * 스프레드시트 ID: YOUR_SPREADSHEET_ID
 */

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID';

// 별점은 1~5 정수로 제한 (비정상 값이 들어와도 화면이 멈추지 않도록)
function normalizeScore(value) {
  const n = Math.round(Number(value));
  if (!n) return 5;
  return Math.min(5, Math.max(1, n));
}

function getTargetSheet() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {}
  if (!ss) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  const sheet = ss.getSheets()[0];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['분류', '책 제목', '작가', '별점', '시작일', '완독일', '표지 이미지', '느낀점']);
    const headerRange = sheet.getRange(1, 1, 1, 8);
    headerRange.setBackground('#064e3b').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function doGet(e) {
  const htmlContent = getHtmlContent();
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('나의 독서 서재')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function getBooks() {
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    const numCols = Math.max(sheet.getLastColumn(), 8);
    const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const books = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const title = row[1] ? String(row[1]).trim() : '';
      if (!title) continue;

      const reviewText = row[7] ? String(row[7]).trim() : (row[6] && !String(row[6]).startsWith('data:image') && !String(row[6]).startsWith('http') ? String(row[6]).trim() : '');
      const firstLine = reviewText.split('\n').map(s => s.trim()).filter(Boolean)[0] || '';

      const formatDate = (val) => {
        if (!val) return '';
        if (val instanceof Date) {
          return Utilities.formatDate(val, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
        }
        return String(val).trim();
      };

      const startDate = formatDate(row[4]);
      const endDate = formatDate(row[5]);

      let coverImage = '';
      if (row[6] && (String(row[6]).startsWith('data:image') || String(row[6]).startsWith('http'))) {
        coverImage = String(row[6]);
      } else if (row[5] && (String(row[5]).startsWith('data:image') || String(row[5]).startsWith('http'))) {
        coverImage = String(row[5]);
      }

      books.push({
        id: i + 2,
        type: row[0] ? String(row[0]).trim() : '기타',
        title: title,
        author: row[2] ? String(row[2]).trim() : '작자 미상',
        score: normalizeScore(row[3]),
        startDate: startDate,
        endDate: endDate,
        imageUrl: coverImage,
        review: reviewText,
        firstLineReview: firstLine
      });
    }

    return books.reverse();
  } catch (err) {
    console.error('도서 목록 조회 오류:', err);
    return [];
  }
}

function getGenres() {
  const defaultGenres = ['소설', 'sf소설', '에세이', '경제/경영', '개발/IT', '인문/철학', '시/문학', '육아', '종교', '기타'];
  try {
    const sheet = getTargetSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const types = sheet.getRange(2, 1, lastRow - 1, 1).getValues()
        .map(r => (r[0] ? String(r[0]).trim() : ''))
        .filter(Boolean);
      return [...new Set([...defaultGenres, ...types])];
    }
  } catch (e) {}
  return defaultGenres;
}

function saveBook(book) {
  if (!book || !book.title || !String(book.title).trim()) {
    throw new Error('책 제목을 입력해주세요.');
  }

  const sheet = getTargetSheet();
  const cleanType = book.type ? String(book.type).trim() : '기타';
  const cleanTitle = String(book.title).trim();
  const cleanAuthor = book.author ? String(book.author).trim() : '';
  const score = normalizeScore(book.score);
  const startDate = book.startDate ? String(book.startDate).trim() : '';
  const endDate = book.endDate ? String(book.endDate).trim() : '';
  const coverImage = book.imageUrl ? String(book.imageUrl).trim() : '';
  const review = book.review ? String(book.review).trim() : '';

  sheet.appendRow([cleanType, cleanTitle, cleanAuthor, score, startDate, endDate, coverImage, review]);

  return {
    success: true,
    message: `'${cleanTitle}' 도서가 서재에 저장되었습니다!`
  };
}

function getHtmlContent() {
  return `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="default">
  <title>나의 독서 서재</title>
  
  <script src="https://cdn.tailwindcss.com"><\/script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/swiper@11/swiper-bundle.min.css" />

  <style>
    @import url('https://fonts.googleapis.com/css2?family=Pretendard:wght@400;500;600;700&family=Nanum+Myeongjo:wght@700&display=swap');
    body {
      font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
      -webkit-tap-highlight-color: transparent;
      user-select: none;
    }
    input, textarea, select { user-select: text; }
    .serif-title { font-family: 'Nanum Myeongjo', serif; }
    .swiper { width: 100%; padding-top: 15px; padding-bottom: 40px; }
    .swiper-slide {
      width: 290px;
      max-width: 82vw;
      transition: all 0.3s ease;
    }
    .swiper-slide-active { transform: scale(1.03); }
    .swiper-pagination-bullet-active {
      background-color: #059669 !important;
      width: 20px !important;
      border-radius: 9999px !important;
    }
    .no-scrollbar::-webkit-scrollbar { display: none; }
    .no-scrollbar { -ms-overflow-style: none; scrollbar-width: none; }
    .drag-over {
      border-color: #059669 !important;
      background-color: #ecfdf5 !important;
    }
  </style>
</head>
<body class="bg-slate-100 min-h-screen text-slate-800 flex flex-col justify-between p-3 sm:p-6">

  <div class="max-w-md w-full mx-auto space-y-4">
    
    <!-- 상단 헤더 -->
    <header class="bg-gradient-to-r from-emerald-600 via-teal-700 to-emerald-800 text-white p-5 rounded-2xl shadow-lg shadow-emerald-700/20 relative overflow-hidden">
      <div class="relative z-10 flex items-center justify-between">
        <div class="flex items-center space-x-3">
          <div class="w-11 h-11 bg-white/20 backdrop-blur-md rounded-xl flex items-center justify-center text-xl text-emerald-100 shadow-inner">
            <i class="fa-solid fa-book-bookmark"></i>
          </div>
          <div>
            <h1 class="text-lg font-bold tracking-tight flex items-center gap-1.5">
              <span>나의 독서 서재</span>
              <span class="text-[10px] bg-white/20 text-emerald-100 font-semibold px-2 py-0.5 rounded-full border border-white/30">2026</span>
            </h1>
            <p class="text-xs text-emerald-100/90 font-medium">읽고 느낀 순간을 기록하는 공간</p>
          </div>
        </div>
        <a href="https://docs.google.com/spreadsheets/d/YOUR_SPREADSHEET_ID/edit" 
           target="_blank" 
           class="px-2.5 py-1.5 bg-white/15 hover:bg-white/25 active:scale-95 text-xs rounded-lg flex items-center gap-1.5 border border-white/20 transition">
          <i class="fa-solid fa-table-cells"></i>
          <span>시트 열기</span>
        </a>
      </div>
      <div class="absolute -right-6 -bottom-8 w-28 h-28 bg-white/10 rounded-full blur-xl pointer-events-none"></div>
    </header>

    <!-- 탭 네비게이션 -->
    <nav class="bg-slate-200/80 p-1 rounded-xl flex text-xs font-bold text-slate-600 shadow-inner">
      <button id="tabWriteBtn" class="flex-1 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 bg-white text-emerald-700 shadow-sm">
        <i class="fa-solid fa-pen-nib"></i>
        <span>새 책 기록하기</span>
      </button>
      <button id="tabListBtn" class="flex-1 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 text-slate-500 hover:text-slate-800">
        <i class="fa-solid fa-layer-group"></i>
        <span>기존 독서 목록 (<span id="bookCountBadge">0</span>)</span>
      </button>
    </nav>

    <!-- ============================================== -->
    <!-- [뷰 1] 새 책 기록하기 (입력 폼) -->
    <!-- ============================================== -->
    <main id="viewWrite" class="space-y-4">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200/80 p-5 space-y-4">
        
        <!-- 분류 (장르) 선택 -->
        <div class="space-y-2">
          <div class="flex items-center justify-between">
            <label class="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <i class="fa-solid fa-folder-open text-emerald-600"></i>
              <span>분류 (장르)</span>
            </label>
            <button type="button" id="toggleGenreInputBtn" class="text-xs font-semibold text-emerald-600 hover:text-emerald-700 active:scale-95 transition flex items-center gap-1">
              <i class="fa-solid fa-plus-circle"></i>
              <span>직접 추가</span>
            </button>
          </div>
          <div id="genreChips" class="flex gap-1.5 overflow-x-auto pb-1 no-scrollbar"></div>
          <div id="customGenreBox" class="hidden pt-1 flex gap-2">
            <input type="text" id="customGenreInput" placeholder="새 장르 입력 (예: 역사소설)" 
                   class="flex-1 px-3 py-2 bg-slate-50 border border-emerald-400 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-emerald-500/20">
            <button type="button" id="addGenreBtn" class="px-3 py-2 bg-emerald-600 text-white rounded-lg text-xs font-bold active:scale-95">추가</button>
          </div>
        </div>

        <!-- 책 제목 & 작가 -->
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <div class="space-y-1">
            <label class="text-xs font-bold text-slate-700">책 제목 <span class="text-rose-500">*</span></label>
            <input type="text" id="bookTitle" placeholder="예: 프로젝트 헤일메리" 
                   class="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-300 focus:border-emerald-600 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20">
          </div>
          <div class="space-y-1">
            <label class="text-xs font-bold text-slate-700">작가</label>
            <input type="text" id="bookAuthor" placeholder="예: 앤디 위어" 
                   class="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-300 focus:border-emerald-600 rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20">
          </div>
        </div>

        <!-- 시작한 날짜 & 완독한 날짜 (나란히 배치) -->
        <div class="grid grid-cols-2 gap-3 bg-slate-50/80 p-3 rounded-xl border border-slate-200/70">
          <div class="space-y-1">
            <label class="text-xs font-bold text-slate-700 flex items-center gap-1">
              <i class="fa-regular fa-calendar-plus text-emerald-600"></i>
              <span>시작한 날짜</span>
            </label>
            <input type="date" id="bookStartDate" 
                   class="w-full px-2.5 py-2 bg-white border border-slate-300 focus:border-emerald-600 rounded-lg text-xs font-medium focus:outline-none">
          </div>
          <div class="space-y-1">
            <label class="text-xs font-bold text-slate-700 flex items-center gap-1">
              <i class="fa-regular fa-calendar-check text-teal-600"></i>
              <span>완독한 날짜</span>
            </label>
            <input type="date" id="bookEndDate" 
                   class="w-full px-2.5 py-2 bg-white border border-slate-300 focus:border-emerald-600 rounded-lg text-xs font-medium focus:outline-none">
          </div>
        </div>

        <!-- 별점 (첫 번째 별 누르면 1개, 두 번째 별 누르면 2개 채워지는 심플 인터랙션) -->
        <div class="bg-slate-50/80 p-3.5 rounded-xl border border-slate-200/70 space-y-1.5">
          <label class="text-xs font-bold text-slate-700 flex items-center gap-1">
            <span>별점</span>
            <span class="text-slate-500 font-medium">(<span id="scoreNumber" class="text-amber-500 font-bold">5</span>점)</span>
          </label>
          
          <!-- 5개 별 터치/클릭 인터랙션 -->
          <div id="starContainer" class="flex items-center gap-2 py-0.5">
            <button type="button" class="star-btn text-2xl transition-transform active:scale-125 focus:outline-none" data-value="1">
              <i class="fa-solid fa-star text-amber-400"></i>
            </button>
            <button type="button" class="star-btn text-2xl transition-transform active:scale-125 focus:outline-none" data-value="2">
              <i class="fa-solid fa-star text-amber-400"></i>
            </button>
            <button type="button" class="star-btn text-2xl transition-transform active:scale-125 focus:outline-none" data-value="3">
              <i class="fa-solid fa-star text-amber-400"></i>
            </button>
            <button type="button" class="star-btn text-2xl transition-transform active:scale-125 focus:outline-none" data-value="4">
              <i class="fa-solid fa-star text-amber-400"></i>
            </button>
            <button type="button" class="star-btn text-2xl transition-transform active:scale-125 focus:outline-none" data-value="5">
              <i class="fa-solid fa-star text-amber-400"></i>
            </button>
          </div>
        </div>

        <!-- 표지 이미지 첨부 (드래그앤드롭 / 캡처 붙여넣기 / 사진첩) -->
        <div class="space-y-1.5">
          <div class="flex items-center justify-between">
            <label class="text-xs font-bold text-slate-700 flex items-center gap-1.5">
              <i class="fa-regular fa-image text-emerald-600"></i>
              <span>표지 사진 첨부</span>
            </label>
            <span class="text-[10px] text-slate-400">캡처 사진 붙여넣기(Ctrl+V) 가능</span>
          </div>

          <div id="imageDropZone" 
               class="border-2 border-dashed border-slate-300 hover:border-emerald-500 rounded-xl p-4 text-center cursor-pointer transition-all bg-slate-50/50 hover:bg-emerald-50/30 relative">
            <input type="file" id="bookImageFileInput" accept="image/*" class="hidden">
            
            <div id="uploadPrompt" class="space-y-1.5 pointer-events-none">
              <div class="w-10 h-10 bg-emerald-100 text-emerald-700 rounded-full flex items-center justify-center mx-auto text-lg">
                <i class="fa-solid fa-cloud-arrow-up"></i>
              </div>
              <p class="text-xs font-semibold text-slate-700">사진을 드래그하거나 터치하여 첨부</p>
              <p class="text-[11px] text-slate-400">캡처 후 바로 붙여넣기(Ctrl+V / Cmd+V)도 됩니다</p>
            </div>

            <div id="imagePreviewContainer" class="hidden flex items-center justify-between gap-3">
              <div class="flex items-center gap-3">
                <img id="imagePreview" src="" alt="표지 미리보기" class="w-14 h-18 object-cover rounded-lg shadow-md border border-slate-200">
                <div class="text-left">
                  <p class="text-xs font-bold text-slate-800" id="imageFileName">표지 이미지 첨부됨</p>
                  <p class="text-[10px] text-emerald-600 font-medium flex items-center gap-1 mt-0.5">
                    <i class="fa-solid fa-circle-check"></i> 준비 완료
                  </p>
                </div>
              </div>
              <button type="button" id="removeImageBtn" class="px-2.5 py-1.5 bg-rose-50 text-rose-600 hover:bg-rose-100 rounded-lg text-xs font-bold active:scale-95 transition">
                <i class="fa-solid fa-trash-can mr-1"></i>삭제
              </button>
            </div>

          </div>
        </div>

        <!-- 느낀점 -->
        <div class="space-y-1">
          <div class="flex items-center justify-between">
            <label class="text-xs font-bold text-slate-700 flex items-center gap-1">
              <i class="fa-solid fa-quote-left text-emerald-600"></i>
              <span>느낀점</span>
            </label>
            <span class="text-[10px] text-emerald-600 font-medium">첫 줄이 카드에 요약 표시됩니다</span>
          </div>
          <textarea id="bookReview" rows="3" placeholder="책을 읽고 느낀 점이나 마음에 남는 문장을 자유롭게 적어보세요..." 
                    class="w-full px-3.5 py-2.5 bg-slate-50 border border-slate-300 focus:border-emerald-600 rounded-xl text-xs focus:outline-none focus:ring-2 focus:ring-emerald-500/20 resize-none"></textarea>
        </div>

        <!-- 저장 버튼 -->
        <button type="button" id="saveBookBtn" 
                class="w-full py-3.5 px-4 bg-emerald-600 hover:bg-emerald-700 active:bg-emerald-800 text-white font-bold rounded-xl shadow-md shadow-emerald-600/30 flex items-center justify-center gap-2 transition active:scale-[0.99] disabled:opacity-50">
          <i class="fa-solid fa-check text-lg"></i>
          <span id="saveBookBtnText" class="text-base tracking-wide">내 서재에 저장하기</span>
        </button>

      </div>
    </main>

    <!-- ============================================== -->
    <!-- [뷰 2] 기존 독서 목록 보기 (3D 슬라이더 & 그리드) -->
    <!-- ============================================== -->
    <section id="viewList" class="hidden space-y-4">
      <div class="flex items-center justify-between px-1">
        <h2 class="text-xs font-bold text-slate-600 flex items-center gap-1.5">
          <i class="fa-solid fa-sparkles text-emerald-600"></i>
          <span>나의 서재 갤러리</span>
        </h2>
        <div class="flex bg-slate-200/80 p-0.5 rounded-lg text-xs font-semibold">
          <button id="modeSliderBtn" class="px-2.5 py-1 rounded-md bg-white text-emerald-700 shadow-sm transition">
            <i class="fa-solid fa-clone mr-1"></i>3D 슬라이드
          </button>
          <button id="modeGridBtn" class="px-2.5 py-1 rounded-md text-slate-500 hover:text-slate-800 transition">
            <i class="fa-solid fa-table-cells-large mr-1"></i>그리드 모아보기
          </button>
        </div>
      </div>

      <div id="listLoading" class="py-16 text-center space-y-3">
        <i class="fa-solid fa-spinner fa-spin text-3xl text-emerald-600"></i>
        <p class="text-xs text-slate-500 font-medium">서재에서 책을 불러오는 중...</p>
      </div>

      <div id="emptyState" class="hidden bg-white rounded-2xl p-10 text-center space-y-3 border border-slate-200">
        <div class="w-14 h-14 bg-emerald-50 text-emerald-600 rounded-2xl flex items-center justify-center mx-auto text-2xl">
          <i class="fa-solid fa-book-open"></i>
        </div>
        <h3 class="font-bold text-slate-700 text-sm">아직 등록된 책이 없습니다</h3>
        <p class="text-xs text-slate-400">첫 번째 독서 기록을 남겨보세요!</p>
        <button type="button" id="goToWriteBtn" class="mt-2 px-4 py-2 bg-emerald-600 text-white text-xs font-bold rounded-lg shadow-md hover:bg-emerald-700 active:scale-95 transition">
          책 기록하러 가기
        </button>
      </div>

      <!-- 3D 카드 슬라이더 -->
      <div id="sliderContainer" class="hidden">
        <div class="swiper myBookSwiper">
          <div class="swiper-wrapper" id="swiperWrapper"></div>
          <div class="swiper-pagination mt-4"></div>
        </div>
        <p class="text-center text-[11px] text-slate-400 mt-1">
          <i class="fa-solid fa-arrows-left-right mr-1"></i>좌우로 스와이프하거나 마우스로 드래그해보세요
        </p>
      </div>

      <!-- 그리드 갤러리 -->
      <div id="gridContainer" class="hidden grid grid-cols-2 gap-3"></div>
    </section>

    <footer class="text-center text-[11px] text-slate-400 py-3">
      <p>💡 홈 화면에 추가하면 전용 독서 앱처럼 바로 열립니다.</p>
    </footer>

  </div>

  <!-- 토스트 알림창 -->
  <div id="toast" class="fixed top-5 left-1/2 -translate-x-1/2 z-50 transform -translate-y-20 opacity-0 transition-all duration-300 pointer-events-none max-w-xs w-full px-4">
    <div class="bg-slate-900/90 backdrop-blur text-white text-xs font-medium py-3 px-4 rounded-xl shadow-xl flex items-center gap-2.5 border border-slate-700">
      <i id="toastIcon" class="fa-solid fa-circle-check text-emerald-400 text-sm"></i>
      <span id="toastMsg">저장되었습니다!</span>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/swiper@11/swiper-bundle.min.js"><\/script>

  <script>
    let books = [];
    let genres = ['소설', 'sf소설', '에세이', '경제/경영', '개발/IT', '인문/철학', '시/문학', '육아', '종교', '기타'];
    let selectedGenre = '소설';
    let currentScore = 5;
    let uploadedImageDataUrl = '';
    let currentViewMode = 'slider';
    let swiperInstance = null;

    const tabWriteBtn = document.getElementById('tabWriteBtn');
    const tabListBtn = document.getElementById('tabListBtn');
    const viewWrite = document.getElementById('viewWrite');
    const viewList = document.getElementById('viewList');
    const bookCountBadge = document.getElementById('bookCountBadge');
    
    const genreChips = document.getElementById('genreChips');
    const toggleGenreInputBtn = document.getElementById('toggleGenreInputBtn');
    const customGenreBox = document.getElementById('customGenreBox');
    const customGenreInput = document.getElementById('customGenreInput');
    const addGenreBtn = document.getElementById('addGenreBtn');
    
    const bookTitle = document.getElementById('bookTitle');
    const bookAuthor = document.getElementById('bookAuthor');
    const bookStartDate = document.getElementById('bookStartDate');
    const bookEndDate = document.getElementById('bookEndDate');
    const scoreNumber = document.getElementById('scoreNumber');
    const starContainer = document.getElementById('starContainer');

    const imageDropZone = document.getElementById('imageDropZone');
    const bookImageFileInput = document.getElementById('bookImageFileInput');
    const uploadPrompt = document.getElementById('uploadPrompt');
    const imagePreviewContainer = document.getElementById('imagePreviewContainer');
    const imagePreview = document.getElementById('imagePreview');
    const imageFileName = document.getElementById('imageFileName');
    const removeImageBtn = document.getElementById('removeImageBtn');

    const bookReview = document.getElementById('bookReview');
    const saveBookBtn = document.getElementById('saveBookBtn');
    const saveBookBtnText = document.getElementById('saveBookBtnText');

    const modeSliderBtn = document.getElementById('modeSliderBtn');
    const modeGridBtn = document.getElementById('modeGridBtn');
    const listLoading = document.getElementById('listLoading');
    const emptyState = document.getElementById('emptyState');
    const sliderContainer = document.getElementById('sliderContainer');
    const gridContainer = document.getElementById('gridContainer');
    const swiperWrapper = document.getElementById('swiperWrapper');
    const goToWriteBtn = document.getElementById('goToWriteBtn');

    const toast = document.getElementById('toast');
    const toastMsg = document.getElementById('toastMsg');
    const toastIcon = document.getElementById('toastIcon');

    window.addEventListener('DOMContentLoaded', () => {
      bookEndDate.value = new Date().toISOString().split('T')[0];
      initGenreChips();
      initStarsInteraction();
      setupImageUploadHandlers();
      setupEventListeners();
      fetchGenres();
      fetchBooks();
    });

    function switchTab(tab) {
      if (tab === 'write') {
        viewWrite.classList.remove('hidden');
        viewList.classList.add('hidden');
        tabWriteBtn.className = 'flex-1 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 bg-white text-emerald-700 shadow-sm';
        tabListBtn.className = 'flex-1 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 text-slate-500 hover:text-slate-800';
      } else {
        viewWrite.classList.add('hidden');
        viewList.classList.remove('hidden');
        tabListBtn.className = 'flex-1 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 bg-white text-emerald-700 shadow-sm';
        tabWriteBtn.className = 'flex-1 py-2.5 rounded-lg transition-all flex items-center justify-center gap-1.5 text-slate-500 hover:text-slate-800';
        renderBookViews();
      }
    }

    function initGenreChips() {
      genreChips.innerHTML = '';
      genres.forEach(g => {
        const isSelected = (g === selectedGenre);
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'px-3 py-1.5 rounded-full text-xs font-semibold whitespace-nowrap transition-all active:scale-95 ' +
          (isSelected ? 'bg-emerald-600 text-white shadow-sm shadow-emerald-600/30' : 'bg-slate-100 text-slate-600 hover:bg-slate-200');
        btn.textContent = g;
        btn.onclick = () => {
          selectedGenre = g;
          initGenreChips();
        };
        genreChips.appendChild(btn);
      });
    }

    // ==============================================
    // 심플 5-Star 터치/클릭 인터랙션
    // (첫 번째 누르면 1개만, 두 번째 누르면 2개까지 채워짐)
    // ==============================================
    function initStarsInteraction() {
      const starButtons = starContainer.querySelectorAll('.star-btn');
      starButtons.forEach(btn => {
        btn.addEventListener('click', () => {
          const val = parseInt(btn.getAttribute('data-value'));
          currentScore = val;
          updateStarsDisplay(val);
        });
      });
      updateStarsDisplay(5);
    }

    function updateStarsDisplay(score) {
      scoreNumber.textContent = score;
      const starButtons = starContainer.querySelectorAll('.star-btn');
      starButtons.forEach(btn => {
        const val = parseInt(btn.getAttribute('data-value'));
        const icon = btn.querySelector('i');
        if (val <= score) {
          icon.className = 'fa-solid fa-star text-amber-400';
        } else {
          icon.className = 'fa-solid fa-star text-slate-200';
        }
      });
    }

    function setupImageUploadHandlers() {
      imageDropZone.addEventListener('click', (e) => {
        if (e.target.closest('#removeImageBtn')) return;
        bookImageFileInput.click();
      });

      bookImageFileInput.addEventListener('change', (e) => {
        if (e.target.files && e.target.files[0]) {
          processImageFile(e.target.files[0]);
        }
      });

      ['dragenter', 'dragover'].forEach(eventName => {
        imageDropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          imageDropZone.classList.add('drag-over');
        });
      });

      ['dragleave', 'drop'].forEach(eventName => {
        imageDropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          imageDropZone.classList.remove('drag-over');
        });
      });

      imageDropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        if (dt && dt.files && dt.files[0]) {
          processImageFile(dt.files[0]);
        }
      });

      window.addEventListener('paste', (e) => {
        const items = (e.clipboardData || e.originalEvent.clipboardData).items;
        for (let index in items) {
          const item = items[index];
          if (item.kind === 'file' && item.type.startsWith('image/')) {
            const blob = item.getAsFile();
            processImageFile(blob, '클립보드 캡처 이미지');
            showToast('캡처한 이미지가 첨부되었습니다!', 'success');
            break;
          }
        }
      });

      removeImageBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        clearUploadedImage();
      });
    }

    function processImageFile(file, customName) {
      if (!file || !file.type.startsWith('image/')) {
        showToast('이미지 파일만 첨부할 수 있습니다.', 'error');
        return;
      }

      const reader = new FileReader();
      reader.onload = function(event) {
        const img = new Image();
        img.onload = function() {
          const canvas = document.createElement('canvas');
          const maxDim = 450;
          let width = img.width;
          let height = img.height;

          if (width > height) {
            if (width > maxDim) {
              height = Math.round((height * maxDim) / width);
              width = maxDim;
            }
          } else {
            if (height > maxDim) {
              width = Math.round((width * maxDim) / height);
              height = maxDim;
            }
          }

          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(img, 0, 0, width, height);

          uploadedImageDataUrl = canvas.toDataURL('image/jpeg', 0.75);

          imagePreview.src = uploadedImageDataUrl;
          imageFileName.textContent = customName || (file.name || '캡처 이미지');
          uploadPrompt.classList.add('hidden');
          imagePreviewContainer.classList.remove('hidden');
          imageDropZone.classList.add('border-emerald-400', 'bg-emerald-50/30');
        };
        img.src = event.target.result;
      };
      reader.readAsDataURL(file);
    }

    function clearUploadedImage() {
      uploadedImageDataUrl = '';
      bookImageFileInput.value = '';
      imagePreview.src = '';
      uploadPrompt.classList.remove('hidden');
      imagePreviewContainer.classList.add('hidden');
      imageDropZone.classList.remove('border-emerald-400', 'bg-emerald-50/30');
    }

    // 시트에서 온 텍스트를 innerHTML에 넣기 전 이스케이프 (XSS 방지)
    function escapeHtml(value) {
      const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
      return String(value == null ? '' : value).replace(/[&<>"']/g, ch => map[ch]);
    }

    function renderStarsEmoji(score) {
      let emojis = '';
      for (let i = 0; i < score; i++) emojis += '⭐';
      return '<span class="text-xs tracking-tight">' + emojis + '</span> <span class="text-[10px] text-slate-400 ml-0.5">(' + score + '점)</span>';
    }

    function getGenreBadgeClass(genre) {
      const g = (genre || '').toLowerCase();
      if (g.includes('sf')) return 'bg-teal-50 text-teal-700 border-teal-200';
      if (g.includes('소설')) return 'bg-emerald-50 text-emerald-800 border-emerald-200';
      if (g.includes('에세이')) return 'bg-amber-50 text-amber-800 border-amber-200';
      if (g.includes('개발')) return 'bg-cyan-50 text-cyan-700 border-cyan-200';
      if (g.includes('경제')) return 'bg-emerald-50 text-emerald-700 border-emerald-200';
      if (g.includes('인문')) return 'bg-slate-100 text-slate-700 border-slate-200';
      return 'bg-emerald-50 text-emerald-700 border-emerald-200';
    }

    function getCoverHtml(book, isLarge = false) {
      if (book.imageUrl && (book.imageUrl.startsWith('data:image') || book.imageUrl.startsWith('http'))) {
        return '<div class="w-full ' + (isLarge ? 'h-48' : 'h-32') + ' bg-slate-100 overflow-hidden relative">' +
               '<img src="' + escapeHtml(book.imageUrl) + '" alt="' + escapeHtml(book.title) + '" data-title="' + escapeHtml(book.title) + '" data-type="' + escapeHtml(book.type) + '" class="w-full h-full object-cover object-center" onerror="this.onerror = null; this.parentElement.outerHTML = getDefaultCoverHtml(this.dataset.title, this.dataset.type, ' + (isLarge ? 'true' : 'false') + ');">' +
               '</div>';
      }
      return getDefaultCoverHtml(book.title, book.type, isLarge);
    }

    function getDefaultCoverHtml(title, type, isLarge) {
      let bgGrad = 'from-emerald-800 via-teal-900 to-slate-900';
      const t = (type || '').toLowerCase();
      if (t.includes('sf')) bgGrad = 'from-teal-800 via-emerald-900 to-slate-900';
      else if (t.includes('에세이')) bgGrad = 'from-emerald-700 via-teal-800 to-stone-900';

      return '<div class="w-full ' + (isLarge ? 'h-48' : 'h-32') + ' bg-gradient-to-br ' + bgGrad + ' flex flex-col justify-between p-3.5 text-white relative overflow-hidden">' +
             '<div class="flex justify-between items-center text-[10px] text-white/70 font-semibold">' +
             '<span>BOOK LOG</span><i class="fa-solid fa-bookmark text-white/50"></i></div>' +
             '<div class="my-auto text-center px-1">' +
             '<p class="serif-title font-bold text-sm tracking-tight text-white/95 line-clamp-2 leading-snug">' + escapeHtml(title) + '</p></div>' +
             '<div class="text-[9px] text-white/50 text-right">' + escapeHtml(type || 'BOOK') + '</div>' +
             '</div>';
    }

    function formatPeriodDate(startDate, endDate) {
      if (startDate && endDate) {
        return startDate + ' ~ ' + endDate;
      }
      if (endDate) return '완독 ' + endDate;
      if (startDate) return '시작 ' + startDate;
      return '';
    }

    function fetchGenres() {
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler((res) => {
            if (Array.isArray(res) && res.length > 0) {
              genres = res;
              initGenreChips();
            }
          })
          .getGenres();
      }
    }

    function fetchBooks() {
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler((res) => {
            listLoading.classList.add('hidden');
            books = Array.isArray(res) ? res : [];
            bookCountBadge.textContent = books.length;
            renderBookViews();
          })
          .withFailureHandler((err) => {
            listLoading.classList.add('hidden');
            showToast('목록 로드 실패: ' + err.message, 'error');
          })
          .getBooks();
      }
    }

    function renderBookViews() {
      if (books.length === 0) {
        emptyState.classList.remove('hidden');
        sliderContainer.classList.add('hidden');
        gridContainer.classList.add('hidden');
        return;
      }

      emptyState.classList.add('hidden');
      if (currentViewMode === 'slider') {
        sliderContainer.classList.remove('hidden');
        gridContainer.classList.add('hidden');
        renderSlider();
      } else {
        sliderContainer.classList.add('hidden');
        gridContainer.classList.remove('hidden');
        renderGrid();
      }
    }

    function renderSlider() {
      swiperWrapper.innerHTML = '';
      books.forEach(b => {
        const slide = document.createElement('div');
        slide.className = 'swiper-slide';

        const coverHtml = getCoverHtml(b, true);
        const starsHtml = renderStarsEmoji(b.score);
        const badgeClass = getGenreBadgeClass(b.type);
        const periodStr = formatPeriodDate(b.startDate, b.endDate);
        const firstLine = b.firstLineReview || '등록된 느낀점이 없습니다.';

        slide.innerHTML = 
          '<div class="bg-white rounded-2xl shadow-lg border border-slate-200/90 overflow-hidden flex flex-col transition-all hover:shadow-xl">' +
            coverHtml +
            '<div class="p-4 flex-1 flex flex-col justify-between space-y-3">' +
              '<div class="space-y-1.5">' +
                '<div class="flex items-center justify-between">' +
                  '<span class="text-[11px] font-bold px-2.5 py-0.5 rounded-md border ' + badgeClass + '">' + escapeHtml(b.type) + '</span>' +
                  '<span class="text-[10px] text-slate-400 font-medium">' + escapeHtml(periodStr) + '</span>' +
                '</div>' +
                '<h3 class="font-bold text-slate-900 text-base leading-snug line-clamp-1">' + escapeHtml(b.title) + '</h3>' +
                '<p class="text-xs text-slate-500 font-medium">' + escapeHtml(b.author || '작자 미상') + '</p>' +
              '</div>' +
              '<div class="bg-slate-50 p-2.5 rounded-xl border border-slate-100">' +
                '<p class="text-xs text-slate-700 italic line-clamp-2 leading-relaxed flex items-start gap-1">' +
                  '<i class="fa-solid fa-quote-left text-[9px] text-emerald-500 mt-1 flex-shrink-0"></i>' +
                  '<span>' + escapeHtml(firstLine) + '</span>' +
                '</p>' +
              '</div>' +
              '<div class="pt-1 border-t border-slate-100 flex items-center justify-between">' +
                starsHtml +
              '</div>' +
            '</div>' +
          '</div>';

        swiperWrapper.appendChild(slide);
      });

      if (swiperInstance) {
        swiperInstance.destroy(true, true);
      }

      swiperInstance = new Swiper('.myBookSwiper', {
        effect: 'coverflow',
        grabCursor: true,
        centeredSlides: true,
        slidesPerView: 'auto',
        initialSlide: 0,
        coverflowEffect: {
          rotate: 0,
          stretch: 15,
          depth: 120,
          modifier: 1.8,
          slideShadows: true,
        },
        pagination: {
          el: '.swiper-pagination',
          clickable: true,
        },
        mousewheel: { forceToAxis: true },
        keyboard: { enabled: true },
      });
    }

    function renderGrid() {
      gridContainer.innerHTML = '';
      books.forEach(b => {
        const card = document.createElement('div');
        card.className = 'bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden flex flex-col';

        const coverHtml = getCoverHtml(b, false);
        const badgeClass = getGenreBadgeClass(b.type);

        card.innerHTML = 
          coverHtml +
          '<div class="p-2.5 flex-1 flex flex-col justify-between space-y-1.5">' +
            '<div>' +
              '<div class="flex items-center gap-1 mb-1">' +
                '<span class="text-[9px] font-bold px-1.5 py-0.2 rounded border ' + badgeClass + '">' + escapeHtml(b.type) + '</span>' +
              '</div>' +
              '<h4 class="font-bold text-xs text-slate-800 line-clamp-1 leading-tight">' + escapeHtml(b.title) + '</h4>' +
              '<p class="text-[10px] text-slate-400 line-clamp-1">' + escapeHtml(b.author || '') + '</p>' +
            '</div>' +
            '<div class="pt-1 border-t border-slate-100 flex items-center justify-between">' +
              '<div class="text-[10px]">' +
                '⭐ <span class="text-slate-600 font-bold">' + b.score + '점</span>' +
              '</div>' +
              '<span class="text-[9px] text-slate-400">' + escapeHtml(b.endDate ? b.endDate.substring(5) : (b.startDate ? b.startDate.substring(5) : '')) + '</span>' +
            '</div>' +
          '</div>';

        gridContainer.appendChild(card);
      });
    }

    function setupEventListeners() {
      tabWriteBtn.addEventListener('click', () => switchTab('write'));
      tabListBtn.addEventListener('click', () => switchTab('list'));
      goToWriteBtn.addEventListener('click', () => switchTab('write'));

      modeSliderBtn.addEventListener('click', () => {
        currentViewMode = 'slider';
        modeSliderBtn.className = 'px-2.5 py-1 rounded-md bg-white text-emerald-700 shadow-sm transition';
        modeGridBtn.className = 'px-2.5 py-1 rounded-md text-slate-500 hover:text-slate-800 transition';
        renderBookViews();
      });

      modeGridBtn.addEventListener('click', () => {
        currentViewMode = 'grid';
        modeGridBtn.className = 'px-2.5 py-1 rounded-md bg-white text-emerald-700 shadow-sm transition';
        modeSliderBtn.className = 'px-2.5 py-1 rounded-md text-slate-500 hover:text-slate-800 transition';
        renderBookViews();
      });

      toggleGenreInputBtn.addEventListener('click', () => {
        customGenreBox.classList.toggle('hidden');
        if (!customGenreBox.classList.contains('hidden')) customGenreInput.focus();
      });

      addGenreBtn.addEventListener('click', handleAddGenre);
      customGenreInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); handleAddGenre(); }
      });

      saveBookBtn.addEventListener('click', handleSaveBook);
    }

    function handleAddGenre() {
      const val = customGenreInput.value.trim();
      if (!val) return;
      if (!genres.includes(val)) genres.push(val);
      selectedGenre = val;
      customGenreInput.value = '';
      customGenreBox.classList.add('hidden');
      initGenreChips();
      showToast("'" + val + "' 장르 추가됨", 'success');
    }

    function handleSaveBook() {
      const title = bookTitle.value.trim();
      if (!title) {
        showToast('책 제목을 입력해주세요.', 'error');
        bookTitle.focus();
        return;
      }

      const bookData = {
        type: selectedGenre,
        title: title,
        author: bookAuthor.value.trim(),
        score: currentScore,
        startDate: bookStartDate.value,
        endDate: bookEndDate.value,
        imageUrl: uploadedImageDataUrl,
        review: bookReview.value.trim()
      };

      setLoading(true);

      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler((res) => {
            setLoading(false);
            onSaveSuccess(bookData);
          })
          .withFailureHandler((err) => {
            setLoading(false);
            showToast('저장 실패: ' + err.message, 'error');
          })
          .saveBook(bookData);
      }
    }

    function onSaveSuccess(bookData) {
      if (navigator.vibrate) navigator.vibrate([40, 60, 40]);
      showToast('서재에 책이 저장되었습니다! 📚', 'success');

      const firstLine = bookData.review.split('\\n').map(s => s.trim()).filter(Boolean)[0] || '';
      books.unshift({
        id: Date.now(),
        type: bookData.type,
        title: bookData.title,
        author: bookData.author,
        score: bookData.score,
        startDate: bookData.startDate,
        endDate: bookData.endDate,
        imageUrl: bookData.imageUrl,
        review: bookData.review,
        firstLineReview: firstLine
      });
      bookCountBadge.textContent = books.length;

      bookTitle.value = '';
      bookAuthor.value = '';
      bookStartDate.value = '';
      bookReview.value = '';
      clearUploadedImage();
      currentScore = 5;
      updateStarsDisplay(5);

      setTimeout(() => {
        switchTab('list');
      }, 700);
    }

    function setLoading(isLoading) {
      saveBookBtn.disabled = isLoading;
      if (isLoading) {
        saveBookBtnText.textContent = '서재에 저장 중...';
        saveBookBtn.querySelector('i').className = 'fa-solid fa-spinner fa-spin text-base';
      } else {
        saveBookBtnText.textContent = '내 서재에 저장하기';
        saveBookBtn.querySelector('i').className = 'fa-solid fa-check text-lg';
      }
    }

    let toastTimeout;
    function showToast(message, type = 'success') {
      clearTimeout(toastTimeout);
      toastMsg.textContent = message;
      toastIcon.className = (type === 'success') 
        ? 'fa-solid fa-circle-check text-emerald-400 text-sm' 
        : 'fa-solid fa-triangle-exclamation text-rose-400 text-sm';

      toast.classList.remove('-translate-y-20', 'opacity-0');
      toast.classList.add('translate-y-0', 'opacity-100');

      toastTimeout = setTimeout(() => {
        toast.classList.remove('translate-y-0', 'opacity-100');
        toast.classList.add('-translate-y-20', 'opacity-0');
      }, 2500);
    }
  </script>
</body>
</html>`;
}
