document.addEventListener('DOMContentLoaded', function() {
    // 동적 페이지 분할 기능
    function dynamicPageSplit() {
        const continuedPages = document.querySelectorAll('.page-continued');
        if (!continuedPages.length) return;
        
        continuedPages.forEach(continuedPage => {
            const content = continuedPage.querySelector('.report-content');
            if (!content) return;
            
            // A4 페이지 최대 높이 (cm to pixels: 1cm = 37.795px)
            const maxHeight = 27.7 * 37.795; // 27.7cm (상하 여백 제외)
            const elements = Array.from(content.children);
            
            let currentHeight = 0;
            let pageCount = parseInt(continuedPage.className.match(/page-continued-(\d+)/)?.[1] || '1');
            let elementsToMove = [];
            let splitTable = null;
            
            // 각 요소의 높이를 체크하면서 페이지 초과 여부 확인
            for (let i = 0; i < elements.length; i++) {
                const element = elements[i];
                
                // 테이블인 경우 행 단위로 처리
                if (element.tagName === 'TABLE') {
                    const tableTop = currentHeight;
                    const rows = element.querySelectorAll('tr');
                    let tableHeight = 0;
                    let splitRowIndex = -1;
                    
                    // 각 행의 높이를 체크
                    for (let j = 0; j < rows.length; j++) {
                        const rowHeight = rows[j].offsetHeight;
                        
                        if (tableTop + tableHeight + rowHeight > maxHeight && j > 1) { // 헤더 행은 항상 포함
                            splitRowIndex = j;
                            break;
                        }
                        tableHeight += rowHeight;
                    }
                    
                    // 테이블을 분할해야 하는 경우
                    if (splitRowIndex > 0) {
                        splitTable = {
                            originalTable: element,
                            splitIndex: splitRowIndex,
                            headers: rows[0].cloneNode(true)
                        };
                        currentHeight = tableTop + tableHeight;
                        continue;
                    }
                }
                
                const elementHeight = element.offsetHeight + 15; // margin 포함
                
                if (currentHeight + elementHeight > maxHeight && currentHeight > 0) {
                    // 현재 요소부터 다음 페이지로 이동
                    elementsToMove = elements.slice(i);
                    break;
                }
                
                currentHeight += elementHeight;
            }
            
            // 테이블 분할이 필요한 경우
            if (splitTable) {
                const { originalTable, splitIndex, headers } = splitTable;
                const rows = originalTable.querySelectorAll('tr');
                
                // 새 테이블 생성 (나머지 행들을 위해)
                const newTable = originalTable.cloneNode(false);
                // 원본 테이블의 모든 클래스 복사
                originalTable.classList.forEach(className => {
                    newTable.classList.add(className);
                });
                newTable.classList.add('split-table-continued');
                newTable.appendChild(headers); // 헤더 행 추가
                
                // 분할 지점 이후의 행들을 새 테이블로 이동
                for (let i = splitIndex; i < rows.length; i++) {
                    newTable.appendChild(rows[i]);
                }
                
                // 새 테이블을 elementsToMove 배열의 맨 앞에 추가
                elementsToMove.unshift(newTable);
            }
            
            // 이동할 요소가 있으면 새 페이지 생성
            if (elementsToMove.length > 0) {
                pageCount++;
                const newPage = document.createElement('div');
                newPage.className = `a4-page page-continued page-continued-${pageCount}`;
                const newContent = document.createElement('div');
                newContent.className = 'report-content';
                newPage.appendChild(newContent);
                
                // 요소들을 새 페이지로 이동
                elementsToMove.forEach(element => {
                    newContent.appendChild(element);
                });
                
                // 현재 페이지 다음에 새 페이지 삽입
                continuedPage.parentNode.insertBefore(newPage, continuedPage.nextSibling);
                
                // 새로 생성된 페이지도 검사 (재귀적으로)
                setTimeout(() => dynamicPageSplit(), 100);
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
    
    // 첫 페이지 처리를 위한 함수
    function handleFirstPage() {
        const firstPage = document.querySelector('.page-1');
        if (!firstPage) return;
        
        const content = firstPage.querySelector('.report-content');
        const bottomFixed = firstPage.querySelector('.page-bottom-fixed');
        if (!content || !bottomFixed) return;
        
        // 하단 고정 컨텐츠의 높이
        const bottomFixedHeight = bottomFixed.offsetHeight;
        
        // 첫 페이지 컨텐츠 영역의 최대 높이
        const maxHeight = (27.7 * 37.795) - bottomFixedHeight - 50; // 여백 포함
        
        const elements = Array.from(content.children);
        let currentHeight = 0;
        let elementsToMove = [];
        let splitTable = null;
        
        // 각 요소의 높이를 체크
        for (let i = 0; i < elements.length; i++) {
            const element = elements[i];
            
            // 테이블 처리
            if (element.tagName === 'TABLE') {
                const tableTop = currentHeight;
                const rows = element.querySelectorAll('tr');
                let tableHeight = 0;
                let splitRowIndex = -1;
                
                for (let j = 0; j < rows.length; j++) {
                    const rowHeight = rows[j].offsetHeight;
                    
                    if (tableTop + tableHeight + rowHeight > maxHeight && j > 1) {
                        splitRowIndex = j;
                        break;
                    }
                    tableHeight += rowHeight;
                }
                
                if (splitRowIndex > 0) {
                    splitTable = {
                        originalTable: element,
                        splitIndex: splitRowIndex,
                        headers: rows[0].cloneNode(true)
                    };
                    break;
                }
            }
            
            const elementHeight = element.offsetHeight + 15;
            
            if (currentHeight + elementHeight > maxHeight && currentHeight > 0) {
                elementsToMove = elements.slice(i);
                break;
            }
            
            currentHeight += elementHeight;
        }
        
        // 첫 페이지에서 넘겨야 할 컨텐츠가 있는 경우
        if (elementsToMove.length > 0 || splitTable) {
            const continuedPage = document.querySelector('.page-continued-1');
            if (continuedPage && continuedPage.querySelector('.report-content')) {
                const continuedContent = continuedPage.querySelector('.report-content');
                
                // 테이블 분할이 필요한 경우
                if (splitTable) {
                    const { originalTable, splitIndex, headers } = splitTable;
                    const rows = originalTable.querySelectorAll('tr');
                    
                    const newTable = originalTable.cloneNode(false);
                    // 원본 테이블의 모든 클래스 복사
                    originalTable.classList.forEach(className => {
                        newTable.classList.add(className);
                    });
                    newTable.classList.add('split-table-continued');
                    newTable.appendChild(headers);
                    
                    for (let i = splitIndex; i < rows.length; i++) {
                        newTable.appendChild(rows[i]);
                    }
                    
                    // 기존 컨텐츠의 맨 앞에 추가
                    continuedContent.insertBefore(newTable, continuedContent.firstChild);
                }
                
                // 나머지 요소들 이동
                const existingElements = Array.from(continuedContent.children);
                elementsToMove.forEach(element => {
                    continuedContent.insertBefore(element, existingElements[0] || null);
                });
            }
        }
    }
    
    // 보고서 페이지가 있을 경우에만 실행
    if (document.querySelector('.a4-page')) {
        // 첫 페이지 처리
        handleFirstPage();
        // 동적 페이지 분할 적용
        dynamicPageSplit();
        checkPageOverflow();
        
        // 윈도우 크기 변경 시 다시 체크
        window.addEventListener('resize', checkPageOverflow);
    }
    // 엑셀 파일 업로드 기능
    const uploadBtn = document.getElementById('upload-btn');
    const excelFile = document.getElementById('excel-file');
    
    if (uploadBtn) {
        uploadBtn.addEventListener('click', async function() {
            if (!excelFile.files.length) {
                alert('엑셀 파일을 선택해주세요.');
                return;
            }
            
            const files = Array.from(excelFile.files);
            const totalFiles = files.length;
            let successCount = 0;
            let failedFiles = [];
            let uploadedSpecimenIds = [];
            
            uploadBtn.disabled = true;
            uploadBtn.textContent = `업로드 중... (0/${totalFiles})`;
            
            // 파일들을 순차적으로 업로드
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const formData = new FormData();
                formData.append('file', file);
                
                uploadBtn.textContent = `업로드 중... (${i}/${totalFiles})`;
                
                try {
                    const response = await fetch('/api/upload-excel', {
                        method: 'POST',
                        body: formData
                    });
                    
                    const data = await response.json();
                    
                    if (data.success) {
                        successCount++;
                        uploadedSpecimenIds.push(data.specimen_id);
                    } else {
                        failedFiles.push({
                            filename: file.name,
                            error: data.error
                        });
                    }
                } catch (error) {
                    failedFiles.push({
                        filename: file.name,
                        error: error.toString()
                    });
                }
                
                uploadBtn.textContent = `업로드 중... (${i + 1}/${totalFiles})`;
            }
            
            // 업로드 완료 후 결과 표시
            uploadBtn.disabled = false;
            uploadBtn.textContent = '업로드';
            
            let resultMessage = `전체 ${totalFiles}개 파일 중 ${successCount}개 업로드 성공\n`;
            
            if (uploadedSpecimenIds.length > 0) {
                resultMessage += `\n업로드된 검체 ID: ${uploadedSpecimenIds.join(', ')}\n`;
            }
            
            if (failedFiles.length > 0) {
                resultMessage += '\n실패한 파일들:\n';
                failedFiles.forEach(failed => {
                    resultMessage += `- ${failed.filename}: ${failed.error}\n`;
                });
            }
            
            alert(resultMessage);
            
            // 목록 새로고침
            loadReportList();
            
            // 파일 입력 필드 초기화
            excelFile.value = '';
            
            // 파일이 1개일 때만 해당 보고서로 이동
            if (totalFiles === 1 && uploadedSpecimenIds.length === 1) {
                window.location.href = '/generate-report?specimen_id=' + uploadedSpecimenIds[0];
            }
            // 여러 파일 업로드 시에는 현재 페이지에 머물기
        });
    }
    
    // 보고서 목록 불러오기
    function loadReportList() {
        const reportList = document.getElementById('report-list');
        if (!reportList) return;
        
        // 캐시 방지를 위해 timestamp 추가
        fetch(`/api/reports?t=${Date.now()}`)
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    reportList.innerHTML = '';
                    
                    if (data.reports.length === 0) {
                        reportList.innerHTML = '<li>저장된 보고서가 없습니다.</li>';
                        return;
                    }
                    
                    data.reports.forEach(report => {
                        const li = document.createElement('li');
                        li.textContent = `${report.specimen_id} (${report.created_at})`;
                        li.addEventListener('click', function() {
                            window.location.href = `/generate-report?specimen_id=${report.specimen_id}`;
                        });
                        reportList.appendChild(li);
                    });
                } else {
                    reportList.innerHTML = '<li>보고서 목록을 불러오는 데 실패했습니다.</li>';
                }
            })
            .catch(error => {
                reportList.innerHTML = '<li>보고서 목록을 불러오는 중 오류가 발생했습니다.</li>';
                console.error('Error:', error);
            });
    }
    
    // 페이지 로드 시 보고서 목록 불러오기
    if (document.getElementById('report-list')) {
        loadReportList();
    }
    
    // 보고서 인쇄 기능
    const printBtn = document.getElementById('print-report');
    if (printBtn) {
        printBtn.addEventListener('click', function() {
            window.print();
        });
    }
});