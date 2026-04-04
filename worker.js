// worker.js (15만건 파싱 전용 백그라운드 스레드)
importScripts('https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js');

self.onmessage = function(e) {
    const { fileBuffer, branchArray, staffPosition } = e.data;
    
    try {
        const branchSet = new Set(branchArray); // 빠른 검색을 위한 Hash Set
        const wb = XLSX.read(fileBuffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        
        let resultMap = new Map(); // 중복 제거용 Map (Key: 임시이름+연락처)

        // 백업 데이터(기존 작업물)인지, 순수 원시 데이터(MG인사관리)인지 식별
        let isBackup = (rows[0] && rows[0][0] === "임시저장이름" && rows[0][1] === "연락처");

        if (isBackup) {
            // [백업 데이터 렌더링 로직]
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row[0] && !row[1]) continue;
                
                const tempName = String(row[0]).trim();
                const phone = String(row[1]).trim();
                const id = String(row[2]).trim();
                
                const key = tempName + "|" + phone;
                if (!resultMap.has(key)) {
                    resultMap.set(key, { tempName, phone, id, role: "기존백업", selected: false });
                }
            }
        } else {
            // [MG 인사관리 파싱 로직 (AutoIt 로직 완벽 이식)]
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const branchCode = String(row[11]).trim(); // L열: 지점코드
                
                // 해당 GRM의 지점코드에 포함된 인원만 필터링
                if (branchSet.has(branchCode)) {
                    const name = String(row[1]).trim();      // B: 이름
                    const role = String(row[2]).trim();      // C: 권한구분
                    const birth = String(row[4]).trim();     // E: 생년월일
                    const sex = String(row[5]).trim();       // F: 성별
                    const phone = String(row[6]).trim();     // G: 연락처
                    const agency = String(row[8]).trim();    // I: 대리점명
                    const head = String(row[10]).trim();     // K: 본부명
                    const branchName = String(row[12]).trim(); // M: 지점명
                    const empNo = String(row[13]).trim();    // N: 사번

                    // 1) 사번 / 생년월일-성별코드 생성 로직
                    let birth6 = birth;
                    if (birth.length === 8) birth6 = birth.substring(2);
                    else if (birth.length > 6) birth6 = birth.substring(0, 6);

                    let yy = parseInt(birth6.substring(0, 2)) || 0;
                    let isAfter2000 = (yy >= 0 && yy <= 24);
                    let sSex = sex.charAt(0);
                    let isMale = (sSex === '남' || sSex === 'M' || sSex === 'm');
                    
                    let sexCode = 0;
                    if (isAfter2000) sexCode = isMale ? 3 : 4;
                    else sexCode = isMale ? 1 : 2;

                    let agentId = `${empNo} / ${birth6}-${sexCode}`;

                    // 2) 임시저장이름 결합 로직
                    let tempName = "";
                    if (agency !== "" && agency === head) {
                        tempName = `${agency} ${branchName} ${name}`;
                    } else {
                        tempName = `${agency} ${head} ${branchName} ${name}`;
                    }
                    tempName = tempName.replace(/\s+/g, ' ').trim(); // 다중 공백 제거

                    // 3) 권한구분(role) 스태프 꼬리표 로직
                    if (role && role !== '판매인(GA)') {
                        if (staffPosition === 'front') {
                            tempName = `[${role}] ${tempName}`;
                        } else {
                            tempName = `${tempName} [${role}]`;
                        }
                    }

                    // 4) 중복 제거 등록
                    const key = tempName + "|" + phone;
                    if (!resultMap.has(key)) {
                        resultMap.set(key, { tempName, phone, id: agentId, role, selected: false });
                    }
                }
            }
        }

        // 연산 결과 반환
        self.postMessage({ status: 'success', data: Array.from(resultMap.values()) });

    } catch(err) {
        self.postMessage({ status: 'error', message: err.message });
    }
};