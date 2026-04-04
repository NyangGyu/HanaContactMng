// worker.js (v1.1.0 - 15만건 파싱 및 위촉일자 필터링)
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
                const key = tempName + "|" + phone;
                if (!resultMap.has(key)) {
                    resultMap.set(key, { tempName, phone, id, role: "", selected: true }); 
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

                    if (role && role !== '판매인(GA)') {
                        tempName = `[${role}] ${tempName}`;
                    }

                    const key = tempName + "|" + phone;
                    if (!resultMap.has(key)) {
                        resultMap.set(key, { tempName, phone, id: agentId, role, selected: true });
                    }
                }
            }
        }
        self.postMessage({ status: 'success', data: Array.from(resultMap.values()) });

    } catch(err) {
        self.postMessage({ status: 'error', message: err.message });
    }
};