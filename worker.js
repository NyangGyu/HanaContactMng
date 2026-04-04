// worker.js (v0.7.1 - 실시간 진행률 보고 및 스마트 병합 엔진)
importScripts('https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js');

self.onmessage = function(e) {
    const { fileBuffer, branchArray, startDate, endDate } = e.data;
    
    try {
        const branchSet = new Set(branchArray);
        
        // 1단계: 바이너리 해독 (이 구간은 퍼센트를 측정할 수 없어 UI가 대기함)
        const wb = XLSX.read(fileBuffer, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
        
        let resultMap = new Map();
        let isBackup = (rows[0] && rows[0][0] === "임시저장이름" && rows[0][1] === "연락처");
        const totalRows = rows.length;

        // 2단계: 데이터 추출 및 병합 (여기서부터 1,000건 단위로 진행률 보고)
        if (isBackup) {
            for (let i = 1; i < totalRows; i++) {
                const row = rows[i];
                if (!row[0] && !row[1]) continue;
                
                const tempName = String(row[0]).trim();
                const phone = String(row[1]).trim();
                const id = String(row[2]).trim();
                const role = String(row[3] || "").trim();
                
                const key = phone.replace(/\D/g, "") || (tempName + i); 
                if (!resultMap.has(key)) {
                    resultMap.set(key, { tempName, phone, id, role, selected: true, isMerged: false }); 
                }

                // 1,000건마다 진행률 발송 (메인 스레드 과부하 방지)
                if (i % 1000 === 0) self.postMessage({ status: 'progress', current: i, total: totalRows });
            }
        } else {
            for (let i = 1; i < totalRows; i++) {
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
                        resultMap.set(phoneKey, { tempName, phone, id: agentId, role, selected: true, isMerged: false });
                    } else {
                        // 스마트 병합 로직
                        let existing = resultMap.get(phoneKey);
                        
                        if (role && !existing.role.includes(role)) {
                            if (existing.role === '판매인(GA)') existing.role = role;
                            else existing.role += `, ${role}`;
                        }
                        
                        if (agentId) {
                            const parseId = (str) => {
                                const parts = str.split('/');
                                return { emp: parts[0] ? parts[0].trim() : "", birth: parts[1] ? parts[1].trim() : "" };
                            };
                            let curr = parseId(existing.id);
                            let next = parseId(agentId);
                            
                            let empSet = new Set(curr.emp ? curr.emp.split(',').map(s=>s.trim()) : []);
                            if(next.emp) empSet.add(next.emp);
                            
                            let birthSet = new Set(curr.birth ? curr.birth.split(',').map(s=>s.trim()) : []);
                            if(next.birth) birthSet.add(next.birth);
                            
                            existing.id = Array.from(empSet).join(', ') + " / " + Array.from(birthSet).join(', ');
                        }
                        existing.isMerged = true; 
                    }
                }
                
                if (i % 1000 === 0) self.postMessage({ status: 'progress', current: i, total: totalRows });
            }
        }
        
        // 최종 100% 완료 발송 후 데이터 반환
        self.postMessage({ status: 'progress', current: totalRows, total: totalRows });
        self.postMessage({ status: 'success', data: Array.from(resultMap.values()) });

    } catch(err) {
        self.postMessage({ status: 'error', message: err.message });
    }
};