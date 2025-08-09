document.addEventListener('DOMContentLoaded', function() {
    // 엑셀 파일 업로드 기능
    const uploadBtn = document.getElementById('upload-btn');
    const excelFile = document.getElementById('excel-file');
    
    if (uploadBtn) {
        uploadBtn.addEventListener('click', function() {
            if (!excelFile.files.length) {
                alert('엑셀 파일을 선택해주세요.');
                return;
            }
            
            const formData = new FormData();
            formData.append('file', excelFile.files[0]);
            
            uploadBtn.disabled = true;
            uploadBtn.textContent = '업로드 중...';
            
            fetch('/api/upload-excel', {
                method: 'POST',
                body: formData
            })
            .then(response => response.json())
            .then(data => {
                uploadBtn.disabled = false;
                uploadBtn.textContent = '업로드';
                
                if (data.success) {
                    const jsonStatus = data.json_saved ? "JSON 파일 저장 성공" : "JSON 파일 저장 실패";
                    alert(`보고서가 성공적으로 업로드되었습니다. (${jsonStatus})`);
                    loadReportList();
                    window.location.href = '/generate-report?specimen_id=' + data.specimen_id;
                } else {
                    alert('업로드 실패: ' + data.error);
                }
            })
            .catch(error => {
                uploadBtn.disabled = false;
                uploadBtn.textContent = '업로드';
                alert('업로드 중 오류가 발생했습니다: ' + error);
            });
        });
    }
    
    // 보고서 목록 불러오기
    function loadReportList() {
        const reportList = document.getElementById('report-list');
        if (!reportList) return;
        
        fetch('/api/reports')
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
    loadReportList();
    
    // 보고서 인쇄 기능
    const printBtn = document.getElementById('print-report');
    if (printBtn) {
        printBtn.addEventListener('click', function() {
            window.print();
        });
    }
});