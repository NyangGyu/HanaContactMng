/*!
 * admin-core.js — 슈퍼관리자 지점코드 분석 엔진 (공유 모듈)
 *
 * hana.html(브라우저)과 deploy.mjs(Node)가 **같은 로직**을 쓰도록 분리한 모듈.
 * 여기에만 규칙을 두고 양쪽이 이걸 로드한다. 한쪽만 고쳐져 결과가 갈리는 사고를 막는 게 목적.
 *
 * 설계 원칙: 이 파일은 **SheetJS에 의존하지 않는다.**
 *   입력도 출력도 순수 AOA(Array of Arrays)라서, 엑셀 읽기/쓰기는 각 호스트가 알아서 한다.
 *   덕분에 노드에서 스텁 없이 그대로 테스트 가능.
 */
(function (root, factory) {
    if (typeof module === 'object' && module.exports) module.exports = factory();
    else root.AdminCore = factory();
}(typeof self !== 'undefined' ? self : this, function () {
    'use strict';

    // GA정보_전체정보 컬럼 매핑 (Excel 열 → 0-based 인덱스)
    // D=대리점명, F=본부명, G=지점코드, H=지점명, I=상태, AC=지알엠1
    var COL = { agency: 3, hq: 5, branch: 6, branchName: 7, status: 8, grm: 28 };
    var STATUS_OK = '정상';
    var CHANGELOG_HEADER = ['구분', '지점코드', '대리점명', '본부명', '지점명', '비고'];

    function s(v) { return v == null ? '' : String(v).trim(); }

    /**
     * GA정보 신규 파일 × 기존 마스터를 대조해 GRM별 추가/삭제/유지를 산출.
     * @param {Array[]} newRows  GA정보_전체정보 시트의 AOA (1행 = 헤더)
     * @param {Array[]} oldRows  기존 GRM_Data.xlsx의 'GRM별_지점코드' 시트 AOA
     * @returns {{result: Array, skipInfo: {skippedRows:number, stoppedCodes:number}}}
     */
    function analyzeBranches(newRows, oldRows) {
        // ── 1. 신규 파일 인덱싱 ─────────────────────────────────────────
        // ⚠️ I열 '상태'가 '정상'인 행만 GRM 매칭 대상. '영업중지' 등이 섞이면 폐점 지점이
        //    담당에 추가되고, 삭제로 잡혀야 할 지점이 kept로 남아 마스터가 오염된다.
        //    단 allBranchInfo(삭제 항목의 지점명 보강용)는 상태 무관하게 채운다.
        var newGrmMap = new Map();      // grmCode -> [지점info]
        var allBranchInfo = new Map();  // branchCode -> 지점info
        var infoFromNormal = new Set(); // allBranchInfo를 정상 행으로 채운 코드
        var normalCodes = new Set();
        var abnormalCodes = new Map();  // branchCode -> 상태값(비정상 행)
        var skippedRows = 0;

        for (var r = 1; r < newRows.length; r++) {
            var row = newRows[r];
            if (!row) continue;
            var branchCode = s(row[COL.branch]);
            if (!branchCode) continue;

            var status = s(row[COL.status]);
            var grm = s(row[COL.grm]);
            var info = {
                branchCode: branchCode,
                agency: s(row[COL.agency]),
                hq: s(row[COL.hq]),
                branchName: s(row[COL.branchName])
            };
            var isNormal = (status === STATUS_OK);
            if (isNormal) normalCodes.add(branchCode);
            else abnormalCodes.set(branchCode, status || '상태미상');

            if (!allBranchInfo.has(branchCode) || (isNormal && !infoFromNormal.has(branchCode))) {
                allBranchInfo.set(branchCode, info);
                if (isNormal) infoFromNormal.add(branchCode);
            }
            if (!isNormal) { skippedRows++; continue; }
            if (grm) {
                if (!newGrmMap.has(grm)) newGrmMap.set(grm, []);
                newGrmMap.get(grm).push(info);
            }
        }

        // 정상 행이 하나도 없는 코드 = 완전히 영업중지된 지점 → 삭제 사유 표기용
        var stoppedStatusByCode = new Map();
        abnormalCodes.forEach(function (st, bc) {
            if (!normalCodes.has(bc)) stoppedStatusByCode.set(bc, st);
        });

        // ── 2. GRM별 비교 ──────────────────────────────────────────────
        var oldRow1 = oldRows[0] || [];
        var oldRow2 = oldRows[1] || [];
        var result = [];
        for (var c = 0; c < oldRow1.length; c++) {
            var codeRaw = s(oldRow1[c]);
            if (!codeRaw) continue;
            var name = s(oldRow2[c]);
            // 슬래시 멀티 사번(직할사업단 공동담당)은 첫 번째를 대표키로
            var matchCode = codeRaw.split('/')[0].trim();

            var oldCodes = [];
            for (var rr = 2; rr < oldRows.length; rr++) {
                var v = oldRows[rr] && oldRows[rr][c];
                if (s(v) !== '') oldCodes.push(s(v));
            }

            var newEntries = newGrmMap.get(matchCode) || [];
            var newCodes = newEntries.map(function (e) { return e.branchCode; });
            var newCodeSet = new Set(newCodes);
            var oldCodeSet = new Set(oldCodes);

            var kept = oldCodes.filter(function (bc) { return newCodeSet.has(bc); });
            var removed = oldCodes.filter(function (bc) { return !newCodeSet.has(bc); });
            var added = newCodes.filter(function (bc) { return !oldCodeSet.has(bc); });

            var newDetailMap = new Map(newEntries.map(function (e) { return [e.branchCode, e]; }));
            // ⚠️ 같은 지점코드의 info 객체는 allBranchInfo/newGrmMap 사이에서 공유되므로,
            //    GRM별 비고(note)를 독립 부여하려면 반드시 복제해서 쓴다.
            var addedDetails = added.map(function (bc) {
                return Object.assign({}, newDetailMap.get(bc), { note: '' });
            });
            var removedDetails = removed.map(function (bc) {
                var base = allBranchInfo.get(bc) || { branchCode: bc, agency: '-', hq: '-', branchName: '-' };
                return Object.assign({}, base, { note: '' });
            });

            result.push({
                colIndex: c, codeRaw: codeRaw, matchCode: matchCode, name: name,
                oldCodes: oldCodes, newCodes: newCodes,
                kept: kept, added: added, removed: removed,
                addedDetails: addedDetails, removedDetails: removedDetails,
                finalCodes: kept.concat(added)
            });
        }

        // ── 3. 이동/인계 추적 → 비고(note) ──────────────────────────────
        var newOwnerByCode = new Map();
        var oldOwnerByCode = new Map();
        result.forEach(function (r2) {
            r2.newCodes.forEach(function (bc) { if (!newOwnerByCode.has(bc)) newOwnerByCode.set(bc, r2); });
            r2.oldCodes.forEach(function (bc) { if (!oldOwnerByCode.has(bc)) oldOwnerByCode.set(bc, r2); });
        });
        result.forEach(function (r2) {
            r2.removedDetails.forEach(function (x) {
                var owner = newOwnerByCode.get(x.branchCode);
                if (owner && owner.matchCode !== r2.matchCode) x.note = '이동 ▶ ' + owner.name;
                else if (stoppedStatusByCode.has(x.branchCode)) x.note = stoppedStatusByCode.get(x.branchCode);
                else x.note = '';
            });
            r2.addedDetails.forEach(function (x) {
                var prev = oldOwnerByCode.get(x.branchCode);
                x.note = (prev && prev.matchCode !== r2.matchCode)
                    ? ('기존 ' + prev.name + ' ▶ 인계') : '신규 개설';
            });
        });

        return {
            result: result,
            skipInfo: { skippedRows: skippedRows, stoppedCodes: stoppedStatusByCode.size }
        };
    }

    /** colIndex 순(=기존 마스터 컬럼 순) 정렬본 반환 */
    function sortResult(result) {
        return result.slice().sort(function (a, b) { return a.colIndex - b.colIndex; });
    }

    /**
     * 'GRM별_지점코드' 시트 AOA.
     * 1행 사번 / 2행 이름 / 3행~ 컬럼별 [...kept, ...added]
     */
    function buildMasterSheet(result) {
        var rs = sortResult(result);
        var aoa = [
            rs.map(function (r) { return r.codeRaw; }),
            rs.map(function (r) { return r.name; })
        ];
        var maxLen = rs.reduce(function (m, r) { return Math.max(m, r.finalCodes.length); }, 0);
        for (var i = 0; i < maxLen; i++) {
            aoa.push(rs.map(function (r) { return r.finalCodes[i] != null ? r.finalCodes[i] : null; }));
        }
        return aoa;
    }

    /**
     * '{YYMMDD} 변경내역' 시트 AOA + 행별 스타일 메타.
     * 유지(kept)는 의도적으로 기록하지 않음 — 변동분 추적이 목적이고 유지분은 마스터에서 확인 가능.
     * @returns {{aoa: Array[], rowMeta: string[]}} rowMeta: 'grm'|'colhdr'|'added'|'removed'|'blank'
     */
    function buildChangeLogSheet(result) {
        var rs = sortResult(result);
        var aoa = [], rowMeta = [];
        rs.forEach(function (r) {
            var a = r.added.length, d = r.removed.length;
            if (a === 0 && d === 0) return;
            aoa.push(['[' + r.codeRaw + ']', r.name, '추가: ' + a + '건 / 삭제: ' + d + '건']); rowMeta.push('grm');
            aoa.push(CHANGELOG_HEADER.slice()); rowMeta.push('colhdr');
            r.addedDetails.forEach(function (x) {
                aoa.push(['추가', x.branchCode, x.agency || '', x.hq || '', x.branchName || '', x.note || '']);
                rowMeta.push('added');
            });
            r.removedDetails.forEach(function (x) {
                aoa.push(['삭제', x.branchCode, x.agency || '', x.hq || '', x.branchName || '', x.note || '']);
                rowMeta.push('removed');
            });
            aoa.push([]); rowMeta.push('blank');
        });
        if (aoa.length === 0) { aoa.push(['변경 사항이 없습니다.']); rowMeta.push('blank'); }
        return { aoa: aoa, rowMeta: rowMeta };
    }

    /**
     * '상품Data' 시트 AOA. product_map.json({상품군: [상품명...]})을 평탄화.
     * GRM_Data_sample.xlsx 형식 = 헤더 `상품명 | 상품군` + 상품군 등장 순서 유지.
     */
    function buildProductSheet(productMap) {
        var aoa = [['상품명', '상품군']];
        Object.keys(productMap || {}).forEach(function (group) {
            var list = productMap[group];
            if (!Array.isArray(list)) return;
            list.forEach(function (nm) { aoa.push([String(nm), group]); });
        });
        return aoa;
    }

    /** 분석 결과 합계 (요약 출력용) */
    function summarize(result) {
        var t = { grmCount: result.length, old: 0, now: 0, kept: 0, added: 0, removed: 0, unchanged: 0 };
        result.forEach(function (r) {
            t.old += r.oldCodes.length;
            t.now += r.newCodes.length;
            t.kept += r.kept.length;
            t.added += r.added.length;
            t.removed += r.removed.length;
            if (r.added.length === 0 && r.removed.length === 0) t.unchanged++;
        });
        t.finalTotal = result.reduce(function (n, r) { return n + r.finalCodes.length; }, 0);
        return t;
    }

    /** YYMMDD 태그 (변경내역 시트명/파일명용). 인자 생략 시 오늘. */
    function dateTag(d) {
        d = d || new Date();
        var p = function (n) { return String(n).padStart(2, '0'); };
        return String(d.getFullYear()).slice(2) + p(d.getMonth() + 1) + p(d.getDate());
    }

    return {
        COL: COL,
        STATUS_OK: STATUS_OK,
        analyzeBranches: analyzeBranches,
        sortResult: sortResult,
        buildMasterSheet: buildMasterSheet,
        buildChangeLogSheet: buildChangeLogSheet,
        buildProductSheet: buildProductSheet,
        summarize: summarize,
        dateTag: dateTag
    };
}));
