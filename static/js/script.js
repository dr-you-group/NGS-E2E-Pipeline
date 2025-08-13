// POST 요청으로 보고서를 새 탭에서 여는 함수 (URL에 specimen_id 노출 방지)
function redirectToReport(specimenId) {
    const form = document.createElement('form');
    form.method = 'POST';
    form.action = '/generate-report';
    form.target = '_blank';  // 새 탭에서 열기
    
    const input = document.createElement('input');
    input.type = 'hidden';
    input.name = 'specimen_id';
    input.value = specimenId;
    
    form.appendChild(input);
    document.body.appendChild(form);
    form.submit();
    
    // form 제거
    document.body.removeChild(form);
}

document.addEventListener('DOMContentLoaded', function() {
    // 실시간 검색 기능
    const searchInput = document.getElementById('search-input');
    const searchResults = document.getElementById('search-results');
    let searchTimeout;
    
    if (searchInput && searchResults) {
        searchInput.addEventListener('input', function() {
            const query = this.value.trim();
            
            // 디바운스로 요청 뚜기 제한
            clearTimeout(searchTimeout);
            
            if (query.length < 1) {
                hideSearchResults();
                return;
            }
            
            searchTimeout = setTimeout(() => {
                performSearch(query);
            }, 300); // 300ms 딘레이
        });
        
        // 검색창 외부 클릭 시 드롭다운 숨기기
        document.addEventListener('click', function(event) {
            if (!searchInput.contains(event.target) && !searchResults.contains(event.target)) {
                hideSearchResults();
            }
        });
        
        // 검색창 포커스 시 드롭다운 다시 보이기
        searchInput.addEventListener('focus', function() {
            if (this.value.trim().length >= 1 && searchResults.children.length > 0) {
                showSearchResults();
            }
        });
    }
    
    async function performSearch(query) {
        try {
            const response = await fetch(`/api/search?q=${encodeURIComponent(query)}`);
            const data = await response.json();
            
            if (data.success) {
                displaySearchResults(data.results);
            } else {
                displayNoResults();
            }
        } catch (error) {
            console.error('검색 오류:', error);
            displayNoResults();
        }
    }
    
    function displaySearchResults(results) {
        searchResults.innerHTML = '';
        
        if (results.length === 0) {
            displayNoResults();
            return;
        }
        
        results.forEach(result => {
            const item = document.createElement('div');
            item.className = 'search-result-item';
            
            item.innerHTML = `
                <div class="result-main">${result.specimen_id}</div>
                <div class="result-details">
                    <span class="result-detail">원발장기: ${result.원발장기}</span>
                    <span class="result-detail">진단: ${result.진단}</span>
                    <span class="result-detail">사인자: ${result.signed1}</span>
                </div>
            `;
            
            item.addEventListener('click', function() {
                redirectToReport(result.specimen_id);
            });
            
            searchResults.appendChild(item);
        });
        
        showSearchResults();
    }
    
    function displayNoResults() {
        searchResults.innerHTML = '<div class="no-results">검색 결과가 없습니다.</div>';
        showSearchResults();
    }
    
    function showSearchResults() {
        searchResults.classList.add('show');
    }
    
    function hideSearchResults() {
        searchResults.classList.remove('show');
    }
    
    // 파일 업로드 기능
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const uploadProgress = document.getElementById('upload-progress');
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.querySelector('.progress-text');
    
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }
    
    function highlight(e) {
        if (uploadArea) {
            uploadArea.classList.add('dragover');
        }
    }
    
    function unhighlight(e) {
        if (uploadArea) {
            uploadArea.classList.remove('dragover');
        }
    }
    
    function handleDrop(e) {
        const files = Array.from(e.dataTransfer.files).filter(file => 
            file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
        );
        
        if (files.length === 0) {
            alert('엑셀 파일만 업로드 가능합니다.');
            return;
        }
        
        handleFiles(files);
    }
    
    if (uploadArea && fileInput) {
        // 클릭으로 파일 선택
        uploadArea.addEventListener('click', () => {
            fileInput.click();
        });
        
        // 파일 입력 변경 시
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFiles(e.target.files);
            }
        });
        
        // Drag and Drop 이벤트
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, preventDefaults, false);
            document.body.addEventListener(eventName, preventDefaults, false);
        });
        
        ['dragenter', 'dragover'].forEach(eventName => {
            uploadArea.addEventListener(eventName, highlight, false);
        });
        
        ['dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, unhighlight, false);
        });
        
        uploadArea.addEventListener('drop', handleDrop, false);
    }
    
    async function handleFiles(files) {
        const fileArray = Array.from(files);
        const totalFiles = fileArray.length;
        let successCount = 0;
        let failedFiles = [];
        let uploadedSpecimenIds = [];
        
        // 진행률 표시
        uploadProgress.style.display = 'block';
        progressFill.style.width = '0%';
        
        // 파일 목록 표시
        const uploadedFilesDiv = document.createElement('div');
        uploadedFilesDiv.className = 'uploaded-files';
        uploadProgress.appendChild(uploadedFilesDiv);
        
        // 각 파일에 대한 상태 표시
        fileArray.forEach(file => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <span class="file-name">${file.name}</span>
                <span class="file-status processing">대기중...</span>
            `;
            uploadedFilesDiv.appendChild(fileItem);
        });
        
        // 파일들을 순차적으로 업로드
        for (let i = 0; i < fileArray.length; i++) {
            const file = fileArray[i];
            const fileItem = uploadedFilesDiv.children[i];
            const statusSpan = fileItem.querySelector('.file-status');
            
            progressText.textContent = `업로드 중... (${i + 1}/${totalFiles})`;
            statusSpan.textContent = '업로드 중...';
            statusSpan.className = 'file-status processing';
            
            const formData = new FormData();
            formData.append('file', file);
            
            try {
                const response = await fetch('/api/upload-excel', {
                    method: 'POST',
                    body: formData
                });
                
                const data = await response.json();
                
                if (data.success) {
                    successCount++;
                    uploadedSpecimenIds.push(data.specimen_id);
                    statusSpan.textContent = `성공 (${data.specimen_id})`;
                    statusSpan.className = 'file-status success';
                } else {
                    failedFiles.push({
                        filename: file.name,
                        error: data.error
                    });
                    statusSpan.textContent = `실패: ${data.error}`;
                    statusSpan.className = 'file-status error';
                }
            } catch (error) {
                failedFiles.push({
                    filename: file.name,
                    error: error.toString()
                });
                statusSpan.textContent = `오류: ${error.message}`;
                statusSpan.className = 'file-status error';
            }
            
            // 진행률 업데이트
            const progress = ((i + 1) / totalFiles) * 100;
            progressFill.style.width = `${progress}%`;
        }
        
        // 업로드 완료
        progressText.textContent = `업로드 완료! 성공: ${successCount}, 실패: ${failedFiles.length}`;
        
        // 파일 입력 초기화
        fileInput.value = '';
        
        // 3초 후 진행률 숨기기
        setTimeout(() => {
            uploadProgress.style.display = 'none';
            uploadedFilesDiv.remove();
        }, 3000);
    }
    // 동적 페이지 분할 기능 - 개선된 버전
    function dynamicPageSplit() {
        const continuedPages = document.querySelectorAll('.page-continued');
        if (!continuedPages.length) return;
        
        continuedPages.forEach(continuedPage => {
            const content = continuedPage.querySelector('.report-content');
            if (!content) return;
            
            // A4 페이지 최대 높이 계산
            const maxHeight = continuedPage.clientHeight - 100; // 패딩 고려
            const elements = Array.from(content.children);
            
            let currentHeight = 0;
            let pageCount = parseInt(continuedPage.className.match(/page-continued-(\d+)/)?.[1] || '1');
            let elementsToMove = [];
            let tableToSplit = null;
            
            // 각 요소의 높이를 실제로 측정하면서 체크
            for (let i = 0; i < elements.length; i++) {
                const element = elements[i];
                const elementHeight = element.offsetHeight;
                
                // 테이블인 경우 행 단위로 분할 검토
                if (element.tagName === 'TABLE') {
                    const rows = Array.from(element.querySelectorAll('tr'));
                    const headerRow = rows[0];
                    let accumulatedTableHeight = headerRow ? headerRow.offsetHeight : 0;
                    let splitAtRow = -1;
                    
                    // 헤더 이후 각 행을 순서대로 체크
                    for (let j = 1; j < rows.length; j++) {
                        const rowHeight = rows[j].offsetHeight;
                        
                        // 현재 높이 + 헤더 + 지금까지의 행들 + 이번 행이 페이지를 넘는지 체크
                        if (currentHeight + accumulatedTableHeight + rowHeight > maxHeight) {
                            if (j > 1) { // 헤더 + 최소 1개 행은 있어야 함
                                splitAtRow = j;
                                break;
                            }
                        }
                        accumulatedTableHeight += rowHeight;
                    }
                    
                    // 테이블 분할이 필요한 경우
                    if (splitAtRow > 0) {
                        tableToSplit = {
                            originalTable: element,
                            splitRowIndex: splitAtRow,
                            headerRow: headerRow.cloneNode(true)
                        };
                        
                        // 현재 페이지에 남을 행들의 높이만 추가
                        let keepHeight = headerRow.offsetHeight;
                        for (let k = 1; k < splitAtRow; k++) {
                            keepHeight += rows[k].offsetHeight;
                        }
                        currentHeight += keepHeight;
                        
                        // 이후 모든 요소는 다음 페이지로
                        elementsToMove = elements.slice(i + 1);
                        break;
                    }
                }
                
                // 일반 요소 처리
                if (currentHeight + elementHeight > maxHeight && currentHeight > 0) {
                    elementsToMove = elements.slice(i);
                    break;
                }
                
                currentHeight += elementHeight;
            }
            
            // 테이블 분할 실행
            if (tableToSplit) {
                const { originalTable, splitRowIndex, headerRow } = tableToSplit;
                const rows = Array.from(originalTable.querySelectorAll('tr'));
                
                // 새 테이블 생성
                const newTable = originalTable.cloneNode(false);
                
                // 원본 테이블의 모든 클래스와 속성 복사
                Array.from(originalTable.attributes).forEach(attr => {
                    if (attr.name !== 'id') { // id는 중복 방지
                        newTable.setAttribute(attr.name, attr.value);
                    }
                });
                originalTable.classList.forEach(className => {
                    newTable.classList.add(className);
                });
                newTable.classList.add('split-table-continued');
                
                // 헤더 행 추가
                newTable.appendChild(headerRow);
                
                // 분할점 이후의 행들을 새 테이블로 이동
                const rowsToMove = [];
                for (let i = splitRowIndex; i < rows.length; i++) {
                    rowsToMove.push(rows[i]);
                }
                
                rowsToMove.forEach(row => {
                    newTable.appendChild(row);
                });
                
                // elementsToMove 배열의 맨 앞에 새 테이블 추가
                elementsToMove.unshift(newTable);
            }
            
            // 새 페이지 생성이 필요한 경우
            if (elementsToMove.length > 0) {
                pageCount++;
                const newPage = document.createElement('div');
                newPage.className = `a4-page page-continued page-continued-${pageCount}`;
                
                const pageContent = document.createElement('div');
                pageContent.className = 'page-border';
                const reportContent = document.createElement('div');
                reportContent.className = 'report-content';
                
                pageContent.appendChild(reportContent);
                newPage.appendChild(pageContent);
                
                // 요소들을 새 페이지로 이동
                elementsToMove.forEach(element => {
                    reportContent.appendChild(element);
                });
                
                // 현재 페이지 다음에 새 페이지 삽입
                continuedPage.parentNode.insertBefore(newPage, continuedPage.nextSibling);
                
                // 새 페이지도 재귀적으로 검사
                setTimeout(() => {
                    dynamicPageSplit();
                }, 100);
                
                return; // 현재 페이지 처리 완료
            }
        });
    }
    
    // A4 페이지 분할 체크 기능
    function checkPageOverflow() {
        const a4Pages = document.querySelectorAll('.a4-page');
        
        a4Pages.forEach((page, pageIndex) => {
            // 각 페이지 내의 섹션들을 확인
            const sections = page.querySelectorAll('.section');
            let totalHeight = 0;
            const pageHeight = page.clientHeight - 100; // 패딩 고려
            
            sections.forEach((section) => {
                const sectionHeight = section.offsetHeight;
                totalHeight += sectionHeight;
                
                // 페이지 높이를 초과하는 섹션에 대해 경고
                if (totalHeight > pageHeight) {
                    console.warn(`Page ${pageIndex + 1}: 컨텐츠가 A4 페이지 크기를 초과합니다!`);
                }
            });
            
            // 페이지 크기 초과 시 시각적 표시
            if (totalHeight > pageHeight) {
                page.classList.add('overflow-warning');
            } else {
                page.classList.remove('overflow-warning');
            }
        });
    }
    
    // 첫 페이지 처리를 위한 함수 - 순서 유지 버전
    function handleFirstPage() {
        const firstPage = document.querySelector('.page-1');
        if (!firstPage) return;
        
        const content = firstPage.querySelector('.report-content');
        if (!content) return;
        
        // 하단 고정 컨텐츠 확인
        const bottomFixed = firstPage.querySelector('.page-bottom-fixed');
        const bottomFixedHeight = bottomFixed ? bottomFixed.offsetHeight : 0;
        
        // 첫 페이지 컨텐츠 영역의 최대 높이 (최대한 공간 활용)
        const pageHeight = firstPage.clientHeight;
        const maxContentHeight = pageHeight - bottomFixedHeight - 40; // 여백을 최소화 (80px → 40px)
        
        console.log(`첫 페이지 전체 높이: ${pageHeight}px, 하단 고정: ${bottomFixedHeight}px, 사용 가능: ${maxContentHeight}px`);
        
        // 원래 순서대로 모든 children 처리
        const elements = Array.from(content.children);
        let currentHeight = 0;
        let elementsToMove = [];
        
        for (let i = 0; i < elements.length; i++) {
            const element = elements[i];
            const elementHeight = element.offsetHeight;
            
            console.log(`요소 ${i}: ${element.tagName}${element.className ? '.' + element.className : ''} - 높이: ${elementHeight}px, 누적: ${currentHeight}px + ${elementHeight}px = ${currentHeight + elementHeight}px`);
            
            // 처음 4개 요소는 무조건 첫 페이지에 유지 (Variants of clinical significance 포함)
            if (i < 4) {
                currentHeight += elementHeight;
                console.log(`필수 요소 ${i} 첫 페이지에 강제 유지, 누적 높이: ${currentHeight}px`);
                continue;
            }
            
            // 4번째 요소부터는 공간 체크
            if (currentHeight + elementHeight > maxContentHeight - 20) { // 20px 여백만 남김
                console.log(`⚠️ 오버플로우! 요소 ${i}부터 다음 페이지로 이동 (필요: ${currentHeight + elementHeight}px, 사용가능: ${maxContentHeight}px)`);
                
                // 테이블인 경우 행별로 분할 시도
                if (element.tagName === 'TABLE' && currentHeight > 0) {
                    const availableSpace = maxContentHeight - currentHeight - 50; // 50px 여백 남김
                    const result = tryTableSplit(element, availableSpace);
                    if (result.canSplit) {
                        console.log(`✂️ 테이블 분할: ${result.splitRowIndex}번째 행에서 분할`);
                        
                        // 테이블 분할 실행
                        const newTable = splitTableAtRow(element, result.splitRowIndex);
                        elementsToMove.push(newTable);
                        
                        // 분할 후 남은 요소들도 이동
                        elementsToMove.push(...elements.slice(i + 1));
                        break;
                    }
                }
                
                // 테이블 분할이 안 되거나 일반 요소인 경우 전체 이동
                elementsToMove.push(...elements.slice(i));
                break;
            }
            
            currentHeight += elementHeight;
            console.log(`✅ 요소 ${i} 첫 페이지에 유지, 누적 높이: ${currentHeight}px`);
        }
        
        // 다음 페이지로 이동할 요소가 있으면 이동
        if (elementsToMove.length > 0) {
            moveElementsToNextPage(elementsToMove, '.page-continued-1');
            console.log(`${elementsToMove.length}개 요소를 다음 페이지로 이동`);
        }
    }
    
    // 테이블 분할 가능성 채크
    function tryTableSplit(table, availableHeight) {
        const rows = Array.from(table.querySelectorAll('tr'));
        if (rows.length <= 1) return { canSplit: false };
        
        const headerRow = rows[0];
        let accumulatedHeight = headerRow.offsetHeight;
        
        // 헤더 + 최소 1개 행은 들어가야 함
        for (let i = 1; i < rows.length; i++) {
            const rowHeight = rows[i].offsetHeight;
            
            if (accumulatedHeight + rowHeight > availableHeight) {
                if (i > 1) { // 헤더 + 1개 이상 행이 들어가면 분할 가능
                    return { canSplit: true, splitRowIndex: i };
                } else {
                    return { canSplit: false }; // 헤더 + 1행도 안 들어가면 분할 불가
                }
            }
            
            accumulatedHeight += rowHeight;
        }
        
        return { canSplit: false }; // 전체 테이블이 들어가면 분할 불필요
    }
    
    // 테이블 분할 실행
    function splitTableAtRow(originalTable, splitRowIndex) {
        const rows = Array.from(originalTable.querySelectorAll('tr'));
        const headerRow = rows[0];
        
        // 새 테이블 생성
        const newTable = originalTable.cloneNode(false);
        
        // 원본 테이블의 모든 속성 복사
        Array.from(originalTable.attributes).forEach(attr => {
            if (attr.name !== 'id') {
                newTable.setAttribute(attr.name, attr.value);
            }
        });
        originalTable.classList.forEach(className => {
            newTable.classList.add(className);
        });
        newTable.classList.add('split-table-continued');
        
        // 헤더 추가
        newTable.appendChild(headerRow.cloneNode(true));
        
        // 분할점 이후의 행들을 새 테이블로 이동
        for (let i = splitRowIndex; i < rows.length; i++) {
            newTable.appendChild(rows[i]);
        }
        
        return newTable;
    }
    
    // 요소들을 다음 페이지로 이동
    function moveElementsToNextPage(elements, nextPageSelector) {
        let nextPage = document.querySelector(nextPageSelector);
        
        // 다음 페이지가 없으면 생성
        if (!nextPage) {
            nextPage = createNewPage(nextPageSelector);
        }
        
        const nextContent = nextPage.querySelector('.report-content');
        if (!nextContent) return;
        
        // 요소들을 다음 페이지로 이동
        elements.forEach(element => {
            nextContent.appendChild(element);
        });
    }
    
    // 새 페이지 생성
    function createNewPage(pageSelector) {
        const pageClass = pageSelector.replace('.', '');
        const pageNumber = pageClass.includes('continued-') ? 
            pageClass.split('continued-')[1] : '1';
        
        const newPage = document.createElement('div');
        newPage.className = `a4-page page-continued ${pageClass}`;
        
        const pageContent = document.createElement('div');
        pageContent.className = 'page-border';
        const reportContent = document.createElement('div');
        reportContent.className = 'report-content';
        
        pageContent.appendChild(reportContent);
        newPage.appendChild(pageContent);
        
        // 적절한 위치에 삽입
        const firstPage = document.querySelector('.page-1');
        if (firstPage) {
            firstPage.parentNode.insertBefore(newPage, firstPage.nextSibling);
        }
        
        return newPage;
    }
    
    // 보고서 페이지가 있을 경우에만 실행 - 개선된 버전
    if (document.querySelector('.a4-page')) {
        // DOM이 완전히 렌더링된 후 페이지 분할 처리
        setTimeout(() => {
            // 첫 페이지 처리
            handleFirstPage();
            // 동적 페이지 분할 적용
            dynamicPageSplit();
            // 오버플로우 체크
            checkPageOverflow();
            
            // 추가적인 체크를 위해 조금 더 기다린 후 한 번 더
            setTimeout(() => {
                handleFirstPage();
                dynamicPageSplit();
                checkPageOverflow();
            }, 500);
        }, 100);
        
        // 윈도우 크기 변경 시 다시 체크
        window.addEventListener('resize', () => {
            setTimeout(() => {
                handleFirstPage();
                dynamicPageSplit();
                checkPageOverflow();
            }, 100);
        });
    }
    
    // PDF 다운로드 버튼 기능
    const pdfDownloadBtn = document.getElementById('pdf-download-btn');
    if (pdfDownloadBtn) {
        pdfDownloadBtn.addEventListener('click', function() {
            // 검체 정보를 파일명으로 사용
            const originalTitle = document.title;
            const specimenId = window.specimenId || 'NGS_보고서';
            document.title = specimenId;
            
            // 인쇄 다이얼로그 열기
            window.print();
            
            // 제목 복원
            setTimeout(() => {
                document.title = originalTitle;
            }, 1000);
        });
    }
});