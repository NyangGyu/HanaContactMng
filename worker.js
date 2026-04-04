// worker.js (v1.0.0 - 15만건 파싱 및 위촉일자 필터링)
importScripts('https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js');

self.onmessage = function(e) {
    const { fileBuffer, branchArray, startDate, endDate } = e.data;
    
    try {
        const branchSet = new Set(branchArray);
        const wb = XLSX.read(fileBuffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        // raw: false 옵션으로 엑셀의 날짜 서식을 문자열로 안전하게 변환
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
        
        let resultMap = new Map();

        let isBackup = (rows[0] && rows[0][0] === "임시저장이름" && rows[0][1] === "연락처");

        if (isBackup) {
            // [작업 백업 파일 렌더링 로직]
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row[0] && !row[1]) continue;
                
                const tempName = String(row[0]).trim();
                const phone = String(row[1]).trim();
                const id = String(row[2]).trim();
                // 백업 데이터는 직급 조작이 불필요하므로 role을 빈 문자열로 둠
                const key = tempName + "|" + phone;
                if (!resultMap.has(key)) {
                    resultMap.set(key, { tempName, phone, id, role: "", selected: true }); // 기본 전체선택
                }
            }
        } else {
            // [MG 인사관리 파싱 로직]
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const branchCode = String(row[11]).trim(); // L열: 지점코드
                
                if (branchSet.has(branchCode)) {
                    // --- 위촉일자(O열) 필터링 로직 ---
                    let rawDateStr = String(row[14] || "").replace(/\D/g, ""); // 숫자만 추출 (YYYYMMDD)
                    if (rawDateStr.length >= 8) {
                        rawDateStr = rawDateStr.substring(0, 8);
                        // 시작일/종료일이 설정되어 있고 범위를 벗어나면 스킵
                        if (startDate && rawDateStr < startDate) continue;
                        if (endDate && rawDateStr > endDate) continue;
                    }

                    const name = String(row[1]).trim();      // B: 이름
                    const role = String(row[2]).trim();      // C: 권한구분
                    const birth = String(row[4]).trim();     // E: 생년월일
                    const sex = String(row[5]).trim();       // F: 성별
                    const phone = String(row[6]).trim();     // G: 연락처
                    const agency = String(row[8]).trim();    // I: 대리점명
                    const head = String(row[10]).trim();     // K: 본부명
                    const branchName = String(row[12]).trim(); // M: 지점명
                    const empNo = String(row[13]).trim();    // N: 사번

                    let birth6 = birth;
                    if (birth.length === 8) birth6 = birth.substring(2);
                    else if (birth.length > 6) birth6 = birth.substring(0, 6);

                    let yy = parseInt(birth6.substring(0, 2)) || 0;
                    let isAfter2000 = (yy >= 0 && yy <= 24);
                    let sSex = sex.charAt(0);
                    let isMale = (sSex === '남' || sSex === 'M' || sSex === 'm');
                    
                    let sexCode = isAfter2000 ? (isMale ? 3 : 4) : (isMale ? 1 : 2);
                    let agentId = `${empNo} / ${birth6}-${sexCode}`;

                    let tempName = (agency !== "" && agency === head) ? 
                        `${agency} ${branchName} ${name}` : 
                        `${agency} ${head} ${branchName} ${name}`;
                    tempName = tempName.replace(/\s+/g, ' ').trim();

                    // 초기 직급은 무조건 [이름 앞]에 배치 (이후 UI에서 동적 변경)
                    if (role && role !== '판매인(GA)') {
                        tempName = `[${role}] ${tempName}`;
                    }

                    const key = tempName + "|" + phone;
                    if (!resultMap.has(key)) {
                        resultMap.set(key, { tempName, phone, id: agentId, role, selected: true }); // 기본 전체선택
                    }
                }
            }
        }

        self.postMessage({ status: 'success', data: Array.from(resultMap.values()) });

    } catch(err) {
        self.postMessage({ status: 'error', message: err.message });
    }
};