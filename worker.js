// worker.js (v1.4.1 - 연락처 기반 중복 통합 엔진 탑재)
importScripts('https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js');

self.onmessage = function(e) {
    const { fileBuffer, branchArray, startDate, endDate } = e.data;
    
    try {
        const branchSet = new Set(branchArray);
        const wb = XLSX.read(fileBuffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
        
        let resultMap = new Map();
        let isBackup = (rows[0] && rows[0][0] === "임시저장이름" && rows[0][1] === "연락처");

        if (isBackup) {
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row[0] && !row[1]) continue;
                
                const tempName = String(row[0]).trim();
                const phone = String(row[1]).trim();
                const id = String(row[2]).trim();
                const role = String(row[3] || "").trim();
                
                const key = phone.replace(/\D/g, "") || (tempName + i); // 번호가 없으면 이름으로 대체
                if (!resultMap.has(key)) {
                    resultMap.set(key, { tempName, phone, id, role, selected: true, isMerged: false }); 
                }
            }
        } else {
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                const branchCode = String(row[11]).trim();
                
                if (branchSet.has(branchCode)) {
                    let rawDateStr = String(row[14] || "").replace(/\D/g, "");
                    if (rawDateStr.length >= 8) {
                        rawDateStr = rawDateStr.substring(0, 8);
                        if (startDate && rawDateStr < startDate) continue;
                        if (endDate && rawDateStr > endDate) continue;
                    }

                    const name = String(row[1]).trim();
                    const role = String(row[2]).trim();
                    const birth = String(row[4]).trim();
                    const sex = String(row[5]).trim();
                    const phone = String(row[6]).trim();
                    const agency = String(row[8]).trim();
                    const head = String(row[10]).trim();
                    const branchName = String(row[12]).trim();
                    const empNo = String(row[13]).trim();

                    // 연락처에서 숫자만 추출하여 완벽한 고유 Key로 사용
                    const phoneKey = phone.replace(/\D/g, "");
                    if(!phoneKey) continue; 

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

                    if (!resultMap.has(phoneKey)) {
                        // 최초 등록
                        resultMap.set(phoneKey, { tempName, phone, id: agentId, role, selected: true, isMerged: false });
                    } else {
                        // ★ 중복 연락처 발견 -> 데이터 병합(Merge)
                        let existing = resultMap.get(phoneKey);
                        
                        // 권한구분(직급) 병합
                        if (role && !existing.role.includes(role)) {
                            if (existing.role === '판매인(GA)') existing.role = role; // 스태프가 우선
                            else existing.role += `, ${role}`;
                        }
                        
                        // 사번(메모) 병합
                        if (agentId && !existing.id.includes(agentId)) {
                            existing.id += ` / ${agentId}`;
                        }
                        
                        existing.isMerged = true; // 중복 통합 플래그 켜기
                    }
                }
            }
        }
        self.postMessage({ status: 'success', data: Array.from(resultMap.values()) });

    } catch(err) {
        self.postMessage({ status: 'error', message: err.message });
    }
};