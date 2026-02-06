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

document.addEventListener('DOMContentLoaded', function () {
    // 실시간 검색 기능
    const searchInput = document.getElementById('search-input');
    const searchResults = document.getElementById('search-results');
    let searchTimeout;

    if (searchInput && searchResults) {
        searchInput.addEventListener('input', function () {
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
        document.addEventListener('click', function (event) {
            if (!searchInput.contains(event.target) && !searchResults.contains(event.target)) {
                hideSearchResults();
            }
        });

        // 검색창 포커스 시 드롭다운 다시 보이기
        searchInput.addEventListener('focus', function () {
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
                    <span class="result-detail">판독의: ${result.signed1}</span>
                </div>
            `;

            item.addEventListener('click', function () {
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

    // Clinical Variants 동적 분할 기능 (Robust Debug Mode with Overlay Control)
    function dynamicClinicalPagination() {
        console.log("%c[Pagination Start] dynamicClinicalPagination triggered", "color: green; font-weight: bold; font-size: 14px;");

        const page1 = document.querySelector('.page-1');
        if (!page1) {
            console.error("[Pagination] .page-1 not found!");
            hideLoadingOverlay(); // Fail safe
            return;
        }

        const content = page1.querySelector('.report-content');
        if (!content) {
            console.error("[Pagination] .report-content not found!");
            hideLoadingOverlay(); // Fail safe
            return;
        }

        // 1. 기준점 설정 (Deadline Calculation)
        // .page-bottom-fixed might be outside .report-content, so search in page1
        let deadlineY = 0;
        const fixedContainer = page1.querySelector('.page-bottom-fixed');

        if (fixedContainer) {
            deadlineY = fixedContainer.offsetTop;
            console.log(`[Pagination] Deadline set by .page-bottom-fixed: ${deadlineY}px`);
        } else {
            // Fallback: look for Other Biomarkers inside content if fixed container missing
            const allSections = Array.from(content.querySelectorAll('.section, div'));
            const otherBiomarkersSection = allSections.find(sec => {
                return sec.textContent && sec.textContent.includes('Other Biomarkers') && sec.tagName !== 'SCRIPT';
            });

            if (otherBiomarkersSection) {
                deadlineY = otherBiomarkersSection.offsetTop;
                console.log(`[Pagination] Deadline set by Other Biomarkers Content: ${deadlineY}px`);
            } else {
                deadlineY = page1.clientHeight - 200;
                console.log(`[Pagination] Deadline set by Page Height limit: ${deadlineY}px`);
            }
        }

        // [DEBUG] Visual Line Removed

        // 2. 검사 대상 설정 (Safe Title Search)
        // Search for H3 containing "clinical significance" regardless of structure
        const allH3 = Array.from(content.querySelectorAll('h3, .result-title'));
        const clinicalTitle = allH3.find(el => {
            return el.innerText.includes('clinical significance') || el.innerText.includes('clinical-highlight');
        });

        if (!clinicalTitle) {
            console.error("[Pagination] '1. Variants of clinical significance' Title not found! (Checked H3/Title classes)");
            hideLoadingOverlay(); // Fail safe
            return;
        }

        const parentContainer = clinicalTitle.parentElement;
        const allChildren = Array.from(parentContainer.children);
        const startIndex = allChildren.indexOf(clinicalTitle) + 1;

        console.log(`[Pagination] Found Title: "${clinicalTitle.innerText.trim()}". Start Index: ${startIndex}, Total Children: ${allChildren.length}`);

        let elementsToCheck = [];
        for (let i = startIndex; i < allChildren.length; i++) {
            const currentElement = allChildren[i];

            // Stop logic
            const textContent = currentElement.innerText || "";
            // REMOVED: Stop at Unknown Variants. We MUST include them to move them if previous content overflows.
            // if (currentElement.tagName === 'H3' && textContent.includes('Variants of unknown significance')) {
            //     console.log(`[Pagination] Stop Condition Met: Next Section 'Unknown Variants'`);
            //     break;
            // }
            if (currentElement.classList.contains('result-title') && i > startIndex) {
                // Also allow other result titles to be collected so they move with the flow
                console.log(`[Pagination] (Debug) Found another result title: ${currentElement.className}. Keeping it in flow.`);
                // break; // DON'T BREAK
            }

            if (currentElement.classList.contains('page-bottom-fixed')) {
                console.log(`[Pagination] Stop Condition Met: Fixed Bottom`);
                break;
            }
            elementsToCheck.push(currentElement);
        }

        console.log(`[Pagination] Elements to be checked for overflow: ${elementsToCheck.length}`);

        // 3. 순회하며 침범 여부 검사 및 분할
        // Pass a callback to hide overlay when done
        // [Repeating Header] Pass the initial Clinical Title
        processPage(page1, elementsToCheck, deadlineY, 1, hideLoadingOverlay, 0, clinicalTitle);
    }

    // processPage and other helpers remain identifying logic as v10...
    function processPage(pageElement, elements, limitY, pageNum, onComplete, depth = 0, activeMainTitle = null) {
        if (depth > 20) {
            console.error("[Pagination CRITICAL] Max recursion depth reached. Stopping to prevent infinite loop.");
            if (onComplete) onComplete();
            return;
        }

        if (elements.length === 0) {
            if (onComplete) onComplete();
            return;
        }

        // ... (Logics are same as before, just ensuring header preservation)
        console.log(`%c[ProcessPage] Page ${pageNum} | Limit: ${limitY} | Elements: ${elements.length}`, "color: blue; font-weight: bold;");

        let splitTableResult = null;
        let overflowIndex = -1;

        // Overflow Check
        for (let i = 0; i < elements.length; i++) {
            const el = elements[i];
            const elementBottom = el.offsetTop + el.offsetHeight;

            // [Repeating Header] Update Active Main Title if encountered
            if (el.matches && (el.matches('.result-title'))) {
                activeMainTitle = el;
                console.log(`  [Header Update] Active Main Title updated via loop: "${el.textContent.trim()}"`);
            }

            console.log(`  [Check] Item [${i}]: ${el.tagName}.${el.className} | Bottom: ${elementBottom} | Limit: ${limitY}`);

            if (elementBottom > limitY - 5) {
                console.warn(`  [OVERFLOW DETECTED] Index ${i} (${el.tagName}) flows over limit! Bottom: ${elementBottom}`);

                // el 자체가 TABLE인 경우도 처리 (Unknown Variants의 DOM 구조)
                const isTableElement = el.tagName === 'TABLE';
                const table = isTableElement ? el : el.querySelector('table');
                console.log(`    -> TABLE Detection: el.tagName=${el.tagName}, querySelector('table') result: ${table ? 'FOUND' : 'NULL'}, isTableElement=${isTableElement}`);
                if (table) {
                    // TABLE이 직접 자식인 경우 제목은 별도 요소(이전 형제)이므로 titleHeight=0
                    const title = isTableElement ? null : el.querySelector('h4, .variant-type');
                    const titleHeight = title ? title.offsetHeight : 0;
                    console.log(`    -> Title: ${title ? title.textContent.substring(0, 30) : 'N/A (TABLE is direct child)'}, titleHeight=${titleHeight}`);

                    if (el.offsetTop + titleHeight > limitY - 5) { // Reduced margin
                        console.log("    -> Title overflow. Decision: MOVE WHOLE BLOCK.");
                        if (i === 0) {
                            console.warn("    -> Index 0 Title overflow. FORCING FIT.");
                            if (elements.length > 1) {
                                overflowIndex = 1;
                                break;
                            } else {
                                overflowIndex = -1;
                                break;
                            }
                        }
                        overflowIndex = i;
                        break;
                    }

                    const tbody = table.querySelector('tbody');
                    const rows = tbody ? Array.from(tbody.rows) : Array.from(table.rows).filter(r => r.parentNode.tagName !== 'THEAD');
                    const thead = table.querySelector('thead');
                    const theadHeight = thead ? thead.offsetHeight : (rows.length > 0 ? rows[0].offsetHeight : 30);

                    const availableSpace = limitY - (el.offsetTop + titleHeight);
                    console.log(`    -> Table found. Avail Space: ${availableSpace}, Header H: ${theadHeight}`);

                    if (availableSpace < theadHeight * 2) {
                        console.log("    -> Not enough space for header. Decision: MOVE WHOLE BLOCK.");
                        if (i === 0) {
                            console.warn("    -> Index 0 Header overflow. FORCING FIT.");
                            if (elements.length > 1) {
                                overflowIndex = 1;
                                break;
                            } else {
                                overflowIndex = -1;
                                break;
                            }
                        }
                        overflowIndex = i;
                        break;
                    }

                    let checkY = el.offsetTop + titleHeight + theadHeight;
                    let splitRowIdx = -1;

                    for (let r = 0; r < rows.length; r++) {
                        checkY += rows[r].offsetHeight;
                        if (checkY > limitY) { // No additional margin - maximize rows
                            splitRowIdx = r;
                            console.log(`    -> Row ${r} causes overflow at Y=${checkY}`);
                            break;
                        }
                    }

                    if (splitRowIdx === -1) {
                        console.log("    -> Logic says rows fit, but element flows over? Margin/Padding issue? Decision: MOVE WHOLE BLOCK.");
                        overflowIndex = i;
                        break;
                    }

                    if (splitRowIdx === 0) {
                        console.log("    -> 1st Data Row overflows. Decision: MOVE WHOLE BLOCK.");
                        overflowIndex = i;
                        break;
                    }

                    console.log(`    -> SPLIT POSSIBLE at Row ${splitRowIdx}.`);
                    splitTableResult = {
                        elementIndex: i,
                        splitRowInTbody: splitRowIdx
                    };
                    overflowIndex = i;
                    break;

                } else {
                    console.log("    -> Non-table element. Decision: MOVE WHOLE BLOCK.");
                    if (i === 0) {
                        console.warn("    -> Index 0 Non-Table. FORCING FIT.");
                        if (elements.length > 1) {
                            overflowIndex = 1;
                            break;
                        } else {
                            overflowIndex = -1;
                            break;
                        }
                    }
                    overflowIndex = i;
                    break;
                }
            }
        }

        if (overflowIndex === -1) {
            console.log("  [ProcessPage] No overflow detected. Page fits.");
            if (onComplete) onComplete();
            return;
        }

        let nextElements = [];

        if (splitTableResult) {
            const idx = splitTableResult.elementIndex;
            const splitPoint = splitTableResult.splitRowInTbody;
            const originalEl = elements[idx];

            console.log(`  [Action] Splitting Element [${idx}] at Row ${splitPoint}`);

            // 1. Clone Creation
            const clonedEl = originalEl.cloneNode(true);
            // originalEl 자체가 TABLE인 경우 처리
            const isOrigTableElement = originalEl.tagName === 'TABLE';
            const origTable = isOrigTableElement ? originalEl : originalEl.querySelector('table');
            const clonedTable = isOrigTableElement ? clonedEl : clonedEl.querySelector('table');
            const origTbody = origTable.querySelector('tbody');
            const clonedTbody = clonedTable.querySelector('tbody');

            // thead 존재 여부 확인
            const hasThead = origTable.querySelector('thead') !== null;
            console.log(`    -> hasThead: ${hasThead}`);

            // 2. Column Width Sync
            const origThs = origTable.querySelectorAll('thead th');
            const clonedThs = clonedTable.querySelectorAll('thead th');
            if (origThs.length > 0) {
                origThs.forEach((th, k) => {
                    const w = th.getBoundingClientRect().width;
                    th.style.width = `${w}px`;
                    th.style.minWidth = `${w}px`;
                    if (clonedThs[k]) {
                        clonedThs[k].style.width = `${w}px`;
                        clonedThs[k].style.minWidth = `${w}px`;
                    }
                });
            }

            // 3. Modify Original (splitPoint 이후 행 삭제)
            const origRows = Array.from(origTbody.rows);
            for (let r = splitPoint; r < origRows.length; r++) {
                if (origRows[r] && origRows[r].parentNode === origTbody) {
                    origTbody.removeChild(origRows[r]);
                }
            }

            // 제목 처리: TABLE이 직접 자식인 경우 이전 형제에서 제목 찾기
            let origTitle = originalEl.querySelector('h4, .variant-type');
            if (!origTitle && isOrigTableElement) {
                // 이전 형제 요소들 중에서 H4 찾기
                let prev = originalEl.previousElementSibling;
                while (prev) {
                    if (prev.matches('h4, .variant-type')) {
                        origTitle = prev;
                        break;
                    }
                    prev = prev.previousElementSibling;
                }
            }
            addPageNumberToTitle(origTitle, pageNum, pageNum + 1);

            // 4. Modify Clone (splitPoint 이전 데이터 행 삭제, 헤더 행은 보존)
            const clonedRows = Array.from(clonedTbody.rows);
            // thead가 없으면 첫 번째 행(index 0)이 헤더이므로 보존
            const startRemoveIdx = hasThead ? 0 : 1;
            console.log(`    -> Removing cloned rows from ${startRemoveIdx} to ${splitPoint - 1}`);
            for (let r = startRemoveIdx; r < splitPoint; r++) {
                if (clonedRows[r] && clonedRows[r].parentNode === clonedTbody) {
                    clonedTbody.removeChild(clonedRows[r]);
                }
            }

            // Clone에 제목 추가 (TABLE이 직접 자식인 경우)
            if (isOrigTableElement && origTitle) {
                const newTitle = origTitle.cloneNode(true);
                addPageNumberToTitle(newTitle, pageNum + 1, pageNum + 1);
                nextElements.push(newTitle);
            } else {
                addPageNumberToTitle(clonedEl.querySelector('h4, .variant-type'), pageNum + 1, pageNum + 1);
            }

            nextElements.push(clonedEl);

            // Add remaining elements
            for (let k = idx + 1; k < elements.length; k++) {
                nextElements.push(elements[k]);
            }

        } else {
            console.log(`  [Action] Moving FULL elements starting from index ${overflowIndex}`);
            for (let k = overflowIndex; k < elements.length; k++) {
                console.log(`    -> MOVING: [${k}] ${elements[k].tagName}.${elements[k].className}`);
                nextElements.push(elements[k]);
            }
        }

        if (nextElements.length > 0) {
            console.log(`  [NewPage] Creating Page ${pageNum + 1} with ${nextElements.length} moved elements.`);
            const newPage = createClinicalContinuedPage(pageElement, pageNum + 1);
            const newContent = newPage.querySelector('.report-content');
            console.log(`    -> newContent found: ${newContent ? 'YES' : 'NO'}`);

            // [Repeating Header] Prepend Main Title Logic (Generalized)
            // Condition: 
            // 1. There is an active main title from the previous page.
            // 2. The *first* element moving to the new page is NOT a new Main Title (H2/H3/result-title).
            //    - If the first element IS a new title (e.g., "▣ Comment"), we don't need to repeat the old one.
            //    - If the first element is content (p, table, div, h4), it means the previous section is continuing.

            if (activeMainTitle && nextElements.length > 0) {
                const firstEl = nextElements[0];
                let isNewMainHeader = (firstEl.matches && (firstEl.matches('.result-title') || firstEl.matches('.section-title') || firstEl.tagName === 'H2' || firstEl.tagName === 'H3'));

                // If not a direct header, check if it's a container (DIV) starting with a header
                if (!isNewMainHeader && (firstEl.tagName === 'DIV' || firstEl.tagName === 'SECTION') && firstEl.children.length > 0) {
                    const firstChild = firstEl.firstElementChild;
                    if (firstChild && (firstChild.matches('.result-title') || firstChild.matches('.section-title') || firstChild.tagName === 'H2' || firstChild.tagName === 'H3')) {
                        isNewMainHeader = true;
                        console.log(`  [DEBUG] Found nested header in DIV: ${firstChild.tagName}.${firstChild.className}`);
                    }
                }

                console.log(`[DEBUG] Repeating Header Check:`);
                console.log(`  - Active Title: "${activeMainTitle.textContent.trim()}"`);
                console.log(`  - First Moved Element: <${firstEl.tagName} class="${firstEl.className}">...`);
                console.log(`  - Is New Main Header?: ${isNewMainHeader}`);

                if (!isNewMainHeader) {
                    const clonedTitle = activeMainTitle.cloneNode(true);
                    clonedTitle.removeAttribute('id');
                    console.log(`    -> DECISION: REPEAT. Prepending title.`);
                    newContent.appendChild(clonedTitle);
                } else {
                    console.log(`    -> DECISION: SKIP. New section detected.`);
                }
            }

            nextElements.forEach((el, idx) => {
                console.log(`    -> APPEND to new page: [${idx}] ${el.tagName}.${el.className}`);
                newContent.appendChild(el);
            });

            setTimeout(() => {
                // Clamping Height: Ensure we don't use an expanded container height.
                // A4 (25.4cm) @ 96dpi is ~960px. Safe limit is ~900-950px.
                let border = newPage.querySelector('.page-border');
                let h = border ? border.clientHeight : 0;

                // If height seems abnormally large (expanded by content), force A4 limit
                if (h > 1000 || h < 100) {
                    console.warn(`[Pagination] Page height ${h}px seems abnormal. Forcing safety limit.`);
                    h = 960; // A4 (25.4cm) exactly
                }

                const newLimit = h - 50;
                console.log(`  [Recursive] Triggering check for Page ${pageNum + 1}. Border H: ${h}, Limit: ${newLimit}`);

                const nextChildren = Array.from(newContent.children);
                if (nextChildren.length > 0) {
                    processPage(newPage, nextChildren, newLimit, pageNum + 1, onComplete, depth + 1, activeMainTitle);
                } else {
                    if (onComplete) onComplete();
                }
            }, 100);
        } else {
            // Should be handled by overflowIndex === -1 check, but for safety
            if (onComplete) onComplete();
        }
    }

    // Helper: Overlay Control
    function hideLoadingOverlay() {
        const overlay = document.getElementById('loading-overlay');
        if (overlay) {
            document.body.classList.add('loaded'); // Trigger CSS opacity transition via class
            setTimeout(() => {
                if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
            }, 600); // 0.6s wait (CSS transition is 0.5s)
            console.log("[Pagination] Loading Overlay Hidden.");
        }
    }

    // Helper: Add Numbering (분할 시 임시 마킹, 나중에 후처리)
    function addPageNumberToTitle(titleEl, current, total) {
        if (!titleEl) return;
        // 제목에 data-split-group 속성 추가 (같은 테이블에서 분할된 제목들 그룹화)
        const groupId = titleEl.dataset.splitGroup || titleEl.textContent.trim().replace(/\s*\(\d+\/\d+\)/g, '').substring(0, 20);
        titleEl.dataset.splitGroup = groupId;
        titleEl.dataset.splitIndex = current;
        console.log(`    -> Title marked: group="${groupId}", index=${current}`);
    }

    // Helper: 분할 완료 후 페이지 번호 업데이트
    function finalizePageNumbers() {
        // 같은 그룹의 제목들을 찾아서 (1/N), (2/N), ..., (N/N) 형식으로 업데이트
        const groups = {};
        document.querySelectorAll('[data-split-group]').forEach(el => {
            const group = el.dataset.splitGroup;
            if (!groups[group]) groups[group] = [];
            groups[group].push(el);
        });

        for (const group in groups) {
            const titles = groups[group];
            const total = titles.length;

            // 분할이 없으면 (total == 1) 번호 표시 안함
            if (total === 1) {
                // 기존 번호 제거
                titles[0].innerHTML = titles[0].innerHTML.replace(/\s*<span.*<\/span>/gi, '').replace(/\s*\(\d+\/\d+\)/g, '');
                continue;
            }

            // splitIndex 기준으로 정렬
            titles.sort((a, b) => parseInt(a.dataset.splitIndex) - parseInt(b.dataset.splitIndex));

            titles.forEach((titleEl, idx) => {
                const current = idx + 1;
                let text = titleEl.innerHTML.replace(/\s*<span.*<\/span>/gi, '').replace(/\s*\(\d+\/\d+\)/g, '');
                titleEl.innerHTML = text + ` <span>(${current}/${total})</span>`;
            });
        }
        console.log(`[Pagination] Page numbers finalized for ${Object.keys(groups).length} groups.`);
    }

    // Helper: Create Page
    function createClinicalContinuedPage(prevPage, pageNum) {
        const newPage = document.createElement('div');
        newPage.className = `a4-page page-clinical-continued page-num-${pageNum}`;
        // Important: Ensure the layout matches exactly
        newPage.innerHTML = `<div class="page-border"><div class="report-content"></div></div>`;
        prevPage.parentNode.insertBefore(newPage, prevPage.nextSibling);
        return newPage;
    }

    // Unknown 요소들을 Clinical 마지막 페이지로 이동 (Clinical 오버플로우 시)
    function attachUnknownToClinicalFlow() {
        console.log("%c[Pagination] attachUnknownToClinicalFlow triggered", "color: orange; font-weight: bold; font-size: 14px;");

        // 1. 마지막 Clinical 페이지 찾기
        const allClinicalPages = document.querySelectorAll('.page-1, .page-clinical-continued');
        if (allClinicalPages.length === 0) {
            console.error("[Pagination] No clinical pages found!");
            hideLoadingOverlay();
            return;
        }
        const lastClinicalPage = allClinicalPages[allClinicalPages.length - 1];
        console.log(`[Pagination] Last clinical page: ${lastClinicalPage.className}`);

        // 2. Unknown 페이지에서 요소 추출
        const unknownPage = document.querySelector('.page-continued-1');
        if (!unknownPage) {
            console.error("[Pagination] .page-continued-1 not found!");
            hideLoadingOverlay();
            return;
        }
        const unknownContent = unknownPage.querySelector('.report-content');
        if (!unknownContent) {
            console.error("[Pagination] .report-content not found in .page-continued-1!");
            hideLoadingOverlay();
            return;
        }

        // 3. additional-info 미리 저장 (마지막에 다시 붙일 예정)
        const additionalInfo = unknownContent.querySelector('.additional-info');
        if (additionalInfo) {
            additionalInfo.remove();
            console.log("[Pagination] .additional-info saved for later relocation.");
        }

        // 4. Unknown 요소들을 배열로 추출
        const unknownElements = Array.from(unknownContent.children);
        console.log(`[Pagination] Moving ${unknownElements.length} elements from Unknown page to Clinical flow.`);

        // 5. 마지막 Clinical 페이지의 content로 이동
        const targetContent = lastClinicalPage.querySelector('.report-content');
        if (!targetContent) {
            console.error("[Pagination] .report-content not found in last clinical page!");
            hideLoadingOverlay();
            return;
        }

        unknownElements.forEach(el => {
            targetContent.appendChild(el);
        });

        // 6. 빈 Unknown 페이지 제거
        unknownPage.remove();
        console.log("[Pagination] Empty .page-continued-1 removed.");

        // 7. 이동된 요소들 재-페이지네이션
        // 마지막 Clinical 페이지의 기준점 계산
        const border = lastClinicalPage.querySelector('.page-border');
        let deadlineY = 960 - 50; // Fallback: A4 기준

        if (border) {
            deadlineY = border.clientHeight - 5;
            console.log(`[Pagination] Deadline for merged page: ${deadlineY}px`);
        }

        // 현재 페이지 번호 계산 (Clinical 페이지 수 기준)
        const currentPageNum = allClinicalPages.length;

        // 이동된 요소들만 다시 체크 (기존 Clinical 요소 제외)
        const allChildrenNow = Array.from(targetContent.children);
        console.log(`[Pagination] Re-checking ${allChildrenNow.length} elements in merged page.`);

        // [Repeating Header] Find existing title in the merged context
        // Try to find the last .result-title in the current content to use as starting point
        let currentActiveTitle = null;
        const resultTitles = Array.from(targetContent.querySelectorAll('.result-title'));
        if (resultTitles.length > 0) {
            currentActiveTitle = resultTitles[resultTitles.length - 1];
            console.log(`[Pagination] Found initial active title for merged page: "${currentActiveTitle.textContent.trim()}"`);
        }

        processPage(lastClinicalPage, allChildrenNow, deadlineY, currentPageNum, () => {
            console.log("[Pagination] Merged page pagination complete.");

            // 페이지 번호 후처리
            finalizePageNumbers();

            // additional-info를 마지막 페이지에 추가
            if (additionalInfo) {
                // 마지막 페이지 다시 찾기 (추가 페이지가 생겼을 수 있음)
                const finalPages = document.querySelectorAll('.page-1, .page-clinical-continued');
                const finalLastPage = finalPages[finalPages.length - 1];
                const finalContent = finalLastPage.querySelector('.report-content');

                if (finalContent) {
                    finalContent.appendChild(additionalInfo);
                    console.log(`[Pagination] .additional-info appended to (${finalLastPage.className}).`);
                }
            }

            hideLoadingOverlay();
        }, 0);
    }

    // Unknown Variants 동적 분할 기능
    function dynamicUnknownPagination() {
        console.log("%c[Pagination Start] dynamicUnknownPagination triggered", "color: purple; font-weight: bold; font-size: 14px;");

        const pageUnknown = document.querySelector('.page-continued-1');
        if (!pageUnknown) {
            console.error("[Pagination] .page-continued-1 not found!");
            return;
        }

        const content = pageUnknown.querySelector('.report-content');
        if (!content) {
            console.error("[Pagination] .report-content not found in page-continued-1!");
            return;
        }

        // 1. 기준점 설정 (페이지 테두리 높이 기준)
        // !! 핵심 수정: offsetTop이 아닌 페이지 자체의 높이를 기준으로 함 !!
        const border = pageUnknown.querySelector('.page-border');
        let deadlineY = 960 - 50; // Fallback: A4 기준

        if (border) {
            // 페이지 테두리의 clientHeight에서 margin을 뺀 값 사용
            deadlineY = border.clientHeight - 5; // Minimal margin (5px) to fit more rows
            console.log(`[Pagination-Unknown] Deadline set by page-border height: ${deadlineY}px (border=${border.clientHeight}px)`);
        } else {
            console.log(`[Pagination-Unknown] Deadline set by Fallback: ${deadlineY}px`);
        }

        // 2. additional-info (안내문) 저장 후 제거 - 마지막 페이지로 이동할 예정
        const additionalInfo = content.querySelector('.additional-info');
        if (additionalInfo) {
            additionalInfo.remove();
            console.log(`[Pagination-Unknown] .additional-info removed temporarily for relocation.`);
        }

        // 3. 검사 대상 설정
        const allChildren = Array.from(content.children);
        console.log(`[Pagination-Unknown] Elements to be checked for overflow: ${allChildren.length}`);

        // [Repeating Header] Identify active title
        let activeTitle = null;
        if (allChildren.length > 0) {
            const first = allChildren[0];
            if (first.matches && (first.matches('.result-title') || first.tagName === 'H3')) {
                activeTitle = first;
                console.log(`[Pagination-Unknown] Initial Active Title: "${activeTitle.innerText.trim()}"`);
            }
        }

        // 4. 순회하며 침범 여부 검사 및 분할
        // 완료 후 콜백에서 페이지 번호 후처리 및 additional-info 배치
        processPage(pageUnknown, allChildren, deadlineY, 2, () => {
            console.log(`[Pagination-Unknown] onComplete callback triggered.`);

            // 4-1. 페이지 번호 후처리
            finalizePageNumbers();

            // 4-2. additional-info를 Unknown Variants 섹션의 마지막 페이지에 추가
            if (additionalInfo) {
                // pageUnknown(.page-continued-1)에서 시작해서 nextSibling으로 마지막 관련 페이지 찾기
                let lastUnknownPage = pageUnknown;
                let sibling = pageUnknown.nextElementSibling;

                // page-clinical-continued 또는 page-num-N 클래스를 가진 연속된 페이지 찾기
                while (sibling && sibling.classList.contains('a4-page') &&
                    (sibling.classList.contains('page-clinical-continued') ||
                        sibling.className.includes('page-num-'))) {
                    lastUnknownPage = sibling;
                    sibling = sibling.nextElementSibling;
                }

                const targetContent = lastUnknownPage.querySelector('.report-content');
                if (targetContent) {
                    targetContent.appendChild(additionalInfo);
                    console.log(`[Pagination-Unknown] .additional-info appended to (${lastUnknownPage.className}).`);
                } else {
                    console.error(`[Pagination-Unknown] .report-content not found in target page.`);
                }
            }
        }, 0);
    }

    // Init
    if (document.querySelector('.a4-page')) {
        setTimeout(() => {
            dynamicClinicalPagination();
            setTimeout(() => {
                // 조건 분기: Clinical 오버플로우 여부 확인
                const clinicalOverflowed = document.querySelector('.page-clinical-continued') !== null;
                console.log(`[Pagination] Clinical overflowed: ${clinicalOverflowed}`);

                if (clinicalOverflowed) {
                    // Clinical이 1페이지 초과 → Unknown을 Clinical 직후에 연속 배치
                    attachUnknownToClinicalFlow();
                } else {
                    // Clinical이 1페이지 이내 → 기존 로직 유지 (별도 페이지)
                    dynamicUnknownPagination();
                }
            }, 500);
        }, 500);

        setTimeout(hideLoadingOverlay, 5000);
    }


    // PDF 다운로드 버튼 기능
    const pdfDownloadBtn = document.getElementById('pdf-download-btn');
    if (pdfDownloadBtn) {
        pdfDownloadBtn.addEventListener('click', async function () {
            // Specification이 로드되었는지 확인
            const specContent = document.getElementById('specification-content');
            if (specContent && specContent.innerHTML.trim() === '') {
                // Specification이 아직 로드되지 않았다면 로드
                if (typeof loadSpecification === 'function') {
                    await loadSpecification();
                    // 로드 완료를 위해 약간 대기
                    await new Promise(resolve => setTimeout(resolve, 300));
                }
            }

            // 페이지 재구성 로직 제거 - 화면에 보이는 그대로 출력
            // handleFirstPage(); // 제거
            // dynamicPageSplit(); // 제거

            // 인쇄용 클래스 추가 제거 - 화면과 동일하게 유지
            // document.body.classList.add('printing'); // 제거

            // 검체 정보를 파일명으로 사용
            const originalTitle = document.title;
            const specimenId = window.specimenId || 'NGS_보고서';
            document.title = specimenId;

            // 바로 인쇄 다이얼로그 열기
            window.print();

            // 제목 복원
            document.title = originalTitle;
        });
    }
});

// PPT 다운로드
function downloadPPT(specimenId) {
    if (!specimenId) {
        alert("검체 번호를 찾을 수 없습니다.");
        return;
    }

    // 사용자에게 다운로드 시작 알림 (UX 개선)
    const btn = document.getElementById('ppt-download-btn');
    const originalText = btn.innerText;
    btn.innerText = "생성 중...";
    btn.disabled = true;

    // 가상의 폼(Form)을 만들어 POST 요청 전송 (파일 다운로드 트리거)
    // fetch()보다 form submit 방식이 파일 다운로드 처리에 더 안정적입니다.
    const form = document.createElement('form');
    form.method = 'POST';
    form.action = '/api/download-pptx'; // app.py에 만든 엔드포인트

    // 검체 번호 데이터 추가
    const input = document.createElement('input');
    input.type = 'hidden';
    input.name = 'specimen_id';
    input.value = specimenId;

    form.appendChild(input);
    document.body.appendChild(form);

    // 전송
    form.submit();

    // 폼 제거 및 버튼 상태 복구 (약간의 딜레이 후)
    document.body.removeChild(form);
    setTimeout(() => {
        btn.innerText = originalText;
        btn.disabled = false;
    }, 3000); // 3초 뒤 버튼 활성화
}