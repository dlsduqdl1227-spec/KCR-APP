// ===== KCR 커핑앱 v3.0 Backend =====
// - 총합 점수 기반 순위 (평균 아님)
// - TOP 20 순위
// - Sweetness x2 적용
// - 강도 표시 개선 (괄호 → 강도:N)
// - Process 선택 (Washed/Natural)
// - 역할 표시 (헤드심사위원/심사위원)

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('KCR 커핑앱 v3.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, user-scalable=no');
}

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function logError_(functionName, error) {
  try {
    var ss = getSS_();
    var logSheet = ss.getSheetByName('ErrorLogs');
    
    if (!logSheet) {
      logSheet = ss.insertSheet('ErrorLogs');
      logSheet.appendRow(['타임스탬프', '함수명', '에러 메시지']);
    }
    
    logSheet.appendRow([new Date(), functionName, error.toString()]);
  } catch (e) {
    Logger.log('logError_ failed: ' + e.toString());
  }
}

function initSheets_() {
  var ss = getSS_();
  
  var userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['userID', 'name', 'phoneNumber', 'role', 'team']);
    userSheet.appendRow(['admin', '관리자', '01099998888', 'admin', '']);
    userSheet.appendRow(['user_head1', '헤드심사위원1', '01011112222', 'head1', '']);
    userSheet.appendRow(['user_judge1-1', '심사위원1-1', '01012345678', 'judge1', '']);
    userSheet.appendRow(['user_judge1-2', '심사위원1-2', '01012345679', 'judge1', '']);
    userSheet.appendRow(['user_judge2-1', '심사위원2-1', '01087654321', 'judge2', '']);
    userSheet.appendRow(['user_judge2-2', '심사위원2-2', '01087654322', 'judge2', '']);
  }
  
  var scoreSheet = ss.getSheetByName('Scores');
  if (!scoreSheet) {
    scoreSheet = ss.insertSheet('Scores');
    scoreSheet.appendRow([
      '제출시간', '컵번호', '심사위원ID', '심사위원명', '팀', '모드', 'Process',
      'Flavor점수', 'Flavor강도', 'Flavor코멘트',
      'Aftertaste점수', 'Aftertaste지속성', 'Aftertaste코멘트',
      'Acidity점수', 'Acidity강도', 'Acidity코멘트',
      'Body점수', 'Body강도', 'Body코멘트',
      'Sweetness점수', 'Sweetness강도', 'Sweetness코멘트',
      'Overall점수', 'Overall코멘트',
      '총점', '검수상태'
    ]);
  }
  
  return {success: true};
}

// ===== v2.9: 성능 최적화 - 캐싱 =====
var userCache = {};
var cacheTimestamp = 0;
var CACHE_DURATION = 300000; // 5분

function loginUser(name, phoneNumber) {
  try {
    var ss = getSS_();
    var userSheet = ss.getSheetByName('Users');
    
    if (!userSheet) {
      initSheets_();
      userSheet = ss.getSheetByName('Users');
    }
    
    // 캐시 확인
    var now = Date.now();
    if (now - cacheTimestamp > CACHE_DURATION) {
      userCache = {};
    }
    
    var cacheKey = name + '_' + phoneNumber;
    if (userCache[cacheKey]) {
      return userCache[cacheKey];
    }
    
    var data = userSheet.getDataRange().getValues();
    
    var inputName = String(name).trim();
    var inputPhone = String(phoneNumber).trim().replace(/[\s-]/g, '');
    
    for (var i = 1; i < data.length; i++) {
      var dbName = String(data[i][1]).trim();
      var dbPhone = String(data[i][2]).trim().replace(/[\s-]/g, '');
      var dbRole = String(data[i][3]).trim();  // toLowerCase 제거, 그대로 사용
      var dbTeam = data[i][4] ? String(data[i][4]).trim() : '';
      
      if (dbName === inputName && dbPhone === inputPhone) {
        // 고유 userID 생성: 전화번호 사용
        var uniqueUserID = 'user_' + inputPhone;
        
        var result = { 
          success: true, 
          userID: uniqueUserID,  // ← 전화번호 기반 고유 ID
          name: dbName, 
          role: dbRole,  // judge1, judge2, head1, head2 등 그대로
          team: dbTeam
        };
        
        userCache[cacheKey] = result;
        cacheTimestamp = now;
        
        return result;
      }
    }
    
    return { success: false, message: '등록되지 않은 사용자입니다.' };
    
  } catch (e) {
    logError_('loginUser', e);
    return { success: false, message: '로그인 오류: ' + e.toString() };
  }
}

function batchSubmitEvaluations(payload) {
  var lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(60000)) {
      return { success: false, message: '다른 사용자가 제출 중입니다. 잠시 후 다시 시도하세요.' };
    }
    
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    
    if (!scoreSheet) { 
      initSheets_(); 
      scoreSheet = ss.getSheetByName('Scores'); 
    }
    
    var now = new Date();
    var status = (payload.mode === 'calibration') ? '검수완료' : '미검수';
    var newRows = [];
    
    if (!payload.cups || payload.cups.length === 0) {
      throw new Error('제출할 컵 데이터가 없습니다.');
    }
    
    // ===== 긴급 수정: 중복 체크 제거! 항상 새로 추가만! =====
    var targetJudgeId = String(payload.judgeId).trim();
    var targetJudgeName = String(payload.judgeName).trim();
    var targetTeam = String(payload.team || '').trim();
    var targetMode = String(payload.mode || 'judge').trim();
    var targetProcess = String(payload.process || '').trim();
    
    for (var i = 0; i < payload.cups.length; i++) {
      var cup = payload.cups[i];
      
      var total = Number(cup.flavor || 0) + 
                  Number(cup.aftertaste || 0) + 
                  Number(cup.acidity || 0) + 
                  Number(cup.body || 0) + 
                  (Number(cup.sweetness || 0) * 2) + 
                  Number(cup.overall || 0);
      
      // 항상 새로운 행으로 추가
      newRows.push([
        now, cup.cupNumber, targetJudgeId, targetJudgeName, targetTeam, targetMode, targetProcess,
        cup.flavor, cup.flavorIntensity || '', cup.noteFlavor || '',
        cup.aftertaste, cup.aftertastePersistence || '', cup.noteAftertaste || '',
        cup.acidity, cup.acidityIntensity || '', cup.noteAcidity || '',
        cup.body, cup.bodyIntensity || '', cup.noteBody || '',
        cup.sweetness, cup.sweetnessIntensity || '', cup.noteSweetness || '',
        cup.overall, cup.noteOverall || '',
        total, status
      ]);
    }
    
    // 모든 데이터를 새 행으로 추가
    if (newRows.length > 0) {
      var lastRow = scoreSheet.getLastRow();
      scoreSheet.getRange(lastRow + 1, 1, newRows.length, 26).setValues(newRows);
    }
    
    SpreadsheetApp.flush();
    
    return { success: true, message: payload.cups.length + '개 컵 저장 완료' };
    
  } catch (e) {
    logError_('batchSubmitEvaluations', e);
    return { success: false, message: '제출 실패: ' + e.toString() };
    
  } finally {
    lock.releaseLock();
  }
}

// ===== v3.0 FINAL: 코멘트 검증 완전 재설계 (점수×강도 교차 검증) =====
function checkCommentLogic(attribute, score, intensity, comment, selectedTags) {
  try {
    score = Number(score);
    intensity = Number(intensity);
    comment = String(comment).trim();
    selectedTags = selectedTags || [];
    
    var lowerComment = comment.toLowerCase();
    
    // ===== 1. 기본 검증 =====
    if (!comment || comment.length < 10) {
      return {
        hasIssue: true,
        severity: 'high',
        issue: '코멘트가 너무 짧습니다 (최소 10자)',
        suggestion: '구체적인 특성을 설명해주세요.'
      };
    }
    
    var hangulJamoOnly = /^[ㄱ-ㅎㅏ-ㅣ\s\u318D\u119E\u11A2\u2022\u2025a\u00B7\uFE55]+$/;
    if (hangulJamoOnly.test(comment)) {
      return {
        hasIssue: true,
        severity: 'critical',
        issue: '자음/모음만 입력되었습니다',
        suggestion: '완성된 한글 문장으로 입력하세요.'
      };
    }
    
    var hangulSyllables = comment.match(/[가-힣]/g);
    var totalHangul = hangulSyllables ? hangulSyllables.length : 0;
    
    if (totalHangul < 5) {
      return {
        hasIssue: true,
        severity: 'high',
        issue: '완성된 한글이 너무 적습니다',
        suggestion: '최소 5자 이상의 완성된 한글로 평가를 작성하세요.'
      };
    }
    
    // ===== 2. 강도 표현 사전 정의 =====
    var intensityDict = {
      'very_weak': ['매우 약한', '미약한', '희미한', '거의 없는', '극히 약한'],
      'weak': ['약한', '약하게', '부족한', '약간 존재', '존재하는'],
      'medium': ['중간', '적절한', '보통', '중간 정도', '균형'],
      'strong': ['뚜렷한', '명확한', '분명한', '강한', '충분한'],
      'very_strong': ['매우 강한', '강렬한', '압도적', '지배적', '극도로']
    };
    
    // ===== 3. 코멘트에서 강도 표현 검출 =====
    var detectedIntensity = null;
    var detectedWord = '';
    
    for (var level in intensityDict) {
      var words = intensityDict[level];
      for (var i = 0; i < words.length; i++) {
        if (lowerComment.indexOf(words[i]) > -1) {
          detectedIntensity = level;
          detectedWord = words[i];
          break;
        }
      }
      if (detectedIntensity) break;
    }
    
    // ===== 4. 강도 검증 (핵심!) =====
    if (detectedIntensity) {
      var isValid = false;
      var expectedWords = '';
      
      // 강도 1-2: very_weak만 허용
      if (intensity <= 2) {
        if (detectedIntensity === 'very_weak') isValid = true;
        expectedWords = '"매우 약한", "미약한", "희미한"';
      }
      // 강도 3: weak 또는 very_weak 허용
      else if (intensity === 3) {
        if (detectedIntensity === 'very_weak' || detectedIntensity === 'weak') isValid = true;
        expectedWords = '"약한", "약간 존재", "존재하는"';
      }
      // 강도 4: medium 또는 weak 허용
      else if (intensity === 4) {
        if (detectedIntensity === 'medium' || detectedIntensity === 'weak') isValid = true;
        expectedWords = '"중간", "적절한", "보통"';
      }
      // 강도 5: medium 또는 strong 허용
      else if (intensity === 5) {
        if (detectedIntensity === 'medium' || detectedIntensity === 'strong') isValid = true;
        expectedWords = '"적절한", "뚜렷한", "분명한"';
      }
      // 강도 6: strong 허용
      else if (intensity === 6) {
        if (detectedIntensity === 'strong') isValid = true;
        expectedWords = '"강한", "뚜렷한", "충분한"';
      }
      // 강도 7: strong 또는 very_strong 허용
      else if (intensity === 7) {
        if (detectedIntensity === 'strong' || detectedIntensity === 'very_strong') isValid = true;
        expectedWords = '"매우 강한", "강렬한", "압도적"';
      }
      
      if (!isValid) {
        return {
          hasIssue: true,
          severity: 'critical',
          issue: '강도 ' + intensity + '인데 "' + detectedWord + '" 표현은 부적절합니다',
          suggestion: '강도 ' + intensity + '는 ' + expectedWords + ' 등으로 표현하세요.'
        };
      }
    }
    
    // ===== 5. 점수×강도 교차 검증 =====
    // 낮은 점수(0-2.99)에서 긍정적 강도 표현 사용 금지
    if (score < 3.0) {
      if (detectedIntensity === 'very_strong' || (detectedIntensity === 'strong' && intensity < 6)) {
        return {
          hasIssue: true,
          severity: 'critical',
          issue: '점수 ' + score.toFixed(2) + '점(낮음)인데 "' + detectedWord + '"는 너무 긍정적입니다',
          suggestion: '점수 3.0 미만에서는 "감지되는", "인지되는", "존재하는" 등을 사용하세요.'
        };
      }
    }
    
    // 높은 점수(4.0+)에서 부정적 강도 표현 사용 금지
    if (score >= 4.0) {
      if (detectedIntensity === 'very_weak') {
        return {
          hasIssue: true,
          severity: 'critical',
          issue: '점수 ' + score.toFixed(2) + '점(높음)인데 "' + detectedWord + '"는 너무 부정적입니다',
          suggestion: '점수 4.0 이상에서는 "뚜렷한", "균형잡힌", "은은한" 등을 사용하세요.'
        };
      }
    }
    
    // ===== 6. 점수 구간 7단계 세밀 분류 =====
    var scoreCategory = '';
    if (score >= 4.75) scoreCategory = 'excellent';
    else if (score >= 4.0) scoreCategory = 'good';
    else if (score >= 3.5) scoreCategory = 'slightly_good';
    else if (score >= 3.0) scoreCategory = 'neutral';
    else if (score >= 2.5) scoreCategory = 'slightly_poor';
    else if (score >= 2.0) scoreCategory = 'poor';
    else scoreCategory = 'very_bad';
    
    // ===== 7. 감성 단어 정의 =====
    var excellentWords = ['매우 우수', '훌륭', '탁월', '뛰어난', '인상적', '복합적', '풍부', '완벽', '최고', '극찬', '우수'];
    var goodWords = ['좋은', '좋음', '좋게', '양호', '깔끔', '조화롭', '균형', '적절', '괜찮', '선명', '명확', '긍정'];
    var neutralWords = ['보통', '평범', '무난', '적당', '그럭저럭', '중간'];
    var poorWords = ['부족', '미흡', '약한', '빈약', '단조롭', '심심', '불균형', '아쉬운'];
    var veryBadWords = ['나쁜', '불량', '불쾌', '거슬리', '심각', '최악', '부정적', '부정'];
    var defectWords = ['발효취', '곰팡이', '흙내', '썩', '부패', '시큼', '쓴맛', '떫은', '텁텁', '비린'];
    
    var hasExcellent = false, hasGood = false, hasNeutral = false;
    var hasPoor = false, hasVeryBad = false, hasDefect = false;
    var foundWord = '';
    
    for (var i = 0; i < excellentWords.length; i++) {
      if (lowerComment.indexOf(excellentWords[i]) > -1) {
        hasExcellent = true;
        foundWord = excellentWords[i];
        break;
      }
    }
    
    for (var i = 0; i < goodWords.length; i++) {
      if (lowerComment.indexOf(goodWords[i]) > -1) {
        hasGood = true;
        if (!foundWord) foundWord = goodWords[i];
        break;
      }
    }
    
    for (var i = 0; i < defectWords.length; i++) {
      if (lowerComment.indexOf(defectWords[i]) > -1) {
        hasDefect = true;
        if (!foundWord) foundWord = defectWords[i];
        break;
      }
    }
    
    for (var i = 0; i < veryBadWords.length; i++) {
      if (lowerComment.indexOf(veryBadWords[i]) > -1) {
        hasVeryBad = true;
        if (!foundWord) foundWord = veryBadWords[i];
        break;
      }
    }
    
    for (var i = 0; i < poorWords.length; i++) {
      if (lowerComment.indexOf(poorWords[i]) > -1) {
        hasPoor = true;
        if (!foundWord) foundWord = poorWords[i];
        break;
      }
    }
    
    // ===== 8. 점수-감성 엄격 검증 =====
    
    // 4.75-5.0: 매우 우수
    if (scoreCategory === 'excellent') {
      if (hasPoor || hasVeryBad || hasDefect) {
        return {
          hasIssue: true,
          severity: 'critical',
          issue: '점수 ' + score.toFixed(2) + '점(매우 우수)인데 부정 표현 "' + foundWord + '"',
          suggestion: '4.75점 이상은 부정 표현을 사용할 수 없습니다.'
        };
      }
    }
    
    // 4.0-4.75: 좋음
    else if (scoreCategory === 'good') {
      if (hasPoor || hasVeryBad || hasDefect) {
        return {
          hasIssue: true,
          severity: 'critical',
          issue: '점수 ' + score.toFixed(2) + '점(좋음)인데 부정 표현 "' + foundWord + '"',
          suggestion: '4.0-4.75점은 부정 표현을 사용할 수 없습니다.'
        };
      }
    }
    
    // 0-2.5: 부정적
    else if (scoreCategory === 'slightly_poor' || scoreCategory === 'poor' || scoreCategory === 'very_bad') {
      if (hasExcellent || hasGood) {
        return {
          hasIssue: true,
          severity: 'critical',
          issue: '점수 ' + score.toFixed(2) + '점(낮음)인데 긍정 표현 "' + foundWord + '"',
          suggestion: '2.5점 미만은 긍정 표현을 사용할 수 없습니다.'
        };
      }
    }
    
    // ===== 9. 결함 단어 + 높은 점수 금지 =====
    if (hasDefect && score >= 3.25) {
      return {
        hasIssue: true,
        severity: 'critical',
        issue: '점수 ' + score.toFixed(2) + '점인데 결함 표현 "' + foundWord + '"',
        suggestion: '발효취, 곰팡이 등 결함 표현은 3.0점 이하에서만 사용하세요.'
      };
    }
    
    // ===== 10. 통과 =====
    return {
      hasIssue: false,
      severity: 'none',
      issue: '',
      suggestion: '코멘트가 점수 및 강도와 일치합니다.'
    };
    
  } catch (e) {
    logError_('checkCommentLogic', e);
    return {
      hasIssue: true,
      severity: 'high',
      issue: '검증 중 오류 발생',
      suggestion: e.toString()
    };
  }
}

function getCalibrationCupNumbers() {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    if (!scoreSheet) return [];
    
    var data = scoreSheet.getDataRange().getValues();
    var cupSet = {};
    
    for (var i = 1; i < data.length; i++) {
      var mode = String(data[i][5]).toLowerCase();
      var cupNum = Number(data[i][1]);
      
      if (mode === 'calibration') {
        cupSet[cupNum] = true;
      }
    }
    
    var cups = Object.keys(cupSet).map(function(c) { return Number(c); });
    cups.sort(function(a, b) { return a - b; });
    
    return cups;
  } catch (e) {
    logError_('getCalibrationCupNumbers', e);
    return [];
  }
}

function getCalibrationResultsByCup(cupNum) {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    if (!scoreSheet) return [];
    
    var data = scoreSheet.getDataRange().getValues();
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var mode = String(data[i][5]).toLowerCase();
      var cup = Number(data[i][1]);
      
      if (mode === 'calibration' && cup === Number(cupNum)) {
        results.push({
          judgeName: String(data[i][3]), 
          team: String(data[i][4]),
          flavor: Number(data[i][7]), 
          flavorIntensity: Number(data[i][8]),
          aftertaste: Number(data[i][10]), 
          aftertastePersistence: Number(data[i][11]),
          acidity: Number(data[i][13]), 
          acidityIntensity: Number(data[i][14]),
          body: Number(data[i][16]), 
          bodyIntensity: Number(data[i][17]),
          sweetness: Number(data[i][19]), 
          sweetnessIntensity: Number(data[i][20]),
          overall: Number(data[i][22]),
          notes: { 
            flavor: String(data[i][9]), 
            aftertaste: String(data[i][12]), 
            acidity: String(data[i][15]), 
            body: String(data[i][18]), 
            sweetness: String(data[i][21]), 
            overall: String(data[i][23]) 
          },
          total: Number(data[i][24])
        });
      }
    }
    return results;
  } catch (e) {
    logError_('getCalibrationResultsByCup', e);
    return [];
  }
}

function getMyPendingReviews(judgeId, judgeName) {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    if (!scoreSheet) return [];
    
    var data = scoreSheet.getDataRange().getValues();
    var list = [];
    var seen = {};
    
    // ===== v3.1: judgeName 필터 추가 =====
    var targetJudgeId = String(judgeId).trim();
    var targetJudgeName = String(judgeName || '').trim();
    
    for (var i = 1; i < data.length; i++) {
      var userId = String(data[i][2]).trim();
      var userName = String(data[i][3]).trim();
      var cupNum = Number(data[i][1]);
      var status = String(data[i][25]).trim();
      var mode = String(data[i][5]).toLowerCase().trim();
      
      // judgeId + judgeName 둘 다 매칭 (judgeName이 없으면 judgeId만)
      var userMatch = (userId === targetJudgeId);
      if (targetJudgeName) {
        userMatch = userMatch && (userName === targetJudgeName);
      }
      
      if (userMatch && status === '미검수' && mode === 'judge') {
        var key = String(cupNum);
        if (!seen[key]) {
          list.push({ cup: cupNum, total: Number(data[i][24]) });
          seen[key] = true;
        }
      }
    }
    
    list.sort(function(a, b) { return a.cup - b.cup; });
    return list;
  } catch (e) {
    logError_('getMyPendingReviews', e);
    return [];
  }
}

function getMyScoreDetail(judgeId, cup) {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    if (!scoreSheet) return null;
    
    var data = scoreSheet.getDataRange().getValues();
    
    // ===== 긴급 수정: 모든 매칭 찾고 최신 것 반환 =====
    var matchedRow = null;
    var matchedIndex = -1;
    
    // 역순으로 검색 (최신 데이터가 아래에 있으므로)
    for (var i = data.length - 1; i >= 1; i--) {
      var userId = String(data[i][2]).trim();
      var cupNum = Number(data[i][1]);
      
      // 정확한 매칭 확인
      if (userId === String(judgeId).trim() && cupNum === Number(cup)) {
        matchedRow = data[i];
        matchedIndex = i;
        break; // 최신 것을 찾았으므로 중단
      }
    }
    
    if (!matchedRow) return null;
    
    return {
      flavor: Number(matchedRow[7]), 
      flavorIntensity: Number(matchedRow[8]),
      aftertaste: Number(matchedRow[10]), 
      aftertastePersistence: Number(matchedRow[11]),
      acidity: Number(matchedRow[13]), 
      acidityIntensity: Number(matchedRow[14]),
      body: Number(matchedRow[16]), 
      bodyIntensity: Number(matchedRow[17]),
      sweetness: Number(matchedRow[19]), 
      sweetnessIntensity: Number(matchedRow[20]),
      overall: Number(matchedRow[22]),
      noteFlavor: String(matchedRow[9]), 
      noteAftertaste: String(matchedRow[12]), 
      noteAcidity: String(matchedRow[15]),
      noteBody: String(matchedRow[18]), 
      noteSweetness: String(matchedRow[21]), 
      noteOverall: String(matchedRow[23]),
      total: Number(matchedRow[24]), 
      status: String(matchedRow[25]), 
      rowIndex: matchedIndex + 1
    };
    
  } catch (e) {
    logError_('getMyScoreDetail', e);
    return null;
  }
}

function updateScore(judgeId, cup, updatedData) {
  var lock = LockService.getScriptLock();
  
  try {
    if (!lock.tryLock(60000)) {
      return { success: false, message: '다른 작업이 진행 중입니다. 잠시 후 다시 시도하세요.' };
    }
    
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    if (!scoreSheet) return { success: false, message: '시트 없음' };
    
    var data = scoreSheet.getDataRange().getValues();
    
    // ===== 긴급 수정: 역순 검색으로 최신 데이터 업데이트 =====
    var targetJudgeId = String(judgeId).trim();
    var targetCup = Number(cup);
    var foundRowIndex = -1;
    
    for (var i = data.length - 1; i >= 1; i--) {
      var userId = String(data[i][2]).trim();
      var cupNum = Number(data[i][1]);
      
      if (userId === targetJudgeId && cupNum === targetCup) {
        foundRowIndex = i + 1; // 스프레드시트는 1-based
        break;
      }
    }
    
    if (foundRowIndex === -1) {
      return { success: false, message: '데이터를 찾을 수 없습니다.' };
    }
    
    var total = Number(updatedData.flavor) + Number(updatedData.aftertaste) + 
                Number(updatedData.acidity) + Number(updatedData.body) + 
                (Number(updatedData.sweetness) * 2) + Number(updatedData.overall);
    
    // 컬럼 8부터 25까지 업데이트 (Flavor ~ Total)
    scoreSheet.getRange(foundRowIndex, 8, 1, 18).setValues([[
      updatedData.flavor, updatedData.flavorIntensity || '', updatedData.noteFlavor || '',
      updatedData.aftertaste, updatedData.aftertastePersistence || '', updatedData.noteAftertaste || '',
      updatedData.acidity, updatedData.acidityIntensity || '', updatedData.noteAcidity || '',
      updatedData.body, updatedData.bodyIntensity || '', updatedData.noteBody || '',
      updatedData.sweetness, updatedData.sweetnessIntensity || '', updatedData.noteSweetness || '',
      updatedData.overall, updatedData.noteOverall || '', total
    ]]);
    
    // 컬럼 26: 검수상태
    scoreSheet.getRange(foundRowIndex, 26).setValue('검수완료');
    
    SpreadsheetApp.flush();
    
    return { success: true };
    
  } catch (e) {
    logError_('updateScore', e);
    return { success: false, message: '저장 실패: ' + e.toString() };
    
  } finally {
    lock.releaseLock();
  }
}

function getTop50Rankings() {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    
    if (!scoreSheet) return { success: false, rankings: [] };
    
    var data = scoreSheet.getDataRange().getValues();
    var cupTotals = {};
    
    for (var i = 1; i < data.length; i++) {
      var mode = String(data[i][5]).toLowerCase();
      var cupNum = String(data[i][1]);
      var status = String(data[i][25]);
      var total = Number(data[i][24]);
      var judgeName = String(data[i][3]);
      
      if (mode === 'judge' && status === '검수완료') {
        if (!cupTotals[cupNum]) cupTotals[cupNum] = { total: 0, count: 0, judges: [] };
        cupTotals[cupNum].total += total;
        cupTotals[cupNum].count += 1;
        cupTotals[cupNum].judges.push(judgeName);
      }
    }
    
    var rankings = [];
    for (var cupNum in cupTotals) {
      // 총합 점수 사용 (평균 아님!)
      var sumTotal = cupTotals[cupNum].total;
      rankings.push({
        cup: Number(cupNum), 
        sumTotal: sumTotal,
        judgeCount: cupTotals[cupNum].count,
        judges: cupTotals[cupNum].judges.join(', ')
      });
    }
    
    // 총합 점수로 정렬
    rankings.sort(function(a, b) { return b.sumTotal - a.sumTotal; });
    
    // TOP 20만 반환
    rankings = rankings.slice(0, 20);
    
    return { success: true, rankings: rankings };
  } catch (e) {
    logError_('getTop50Rankings', e);
    return { success: false, rankings: [] };
  }
}

// ===== v3.0: 관리자용 검수 완료 목록 =====
function getAllCompletedReviews() {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    
    if (!scoreSheet) return [];
    
    var data = scoreSheet.getDataRange().getValues();
    var cupTotals = {};
    
    for (var i = 1; i < data.length; i++) {
      var mode = String(data[i][5]).toLowerCase().trim();
      var cupNum = String(data[i][1]).trim();
      var status = String(data[i][25]).trim();
      var total = Number(data[i][24]);
      
      if (mode === 'judge' && status === '검수완료') {
        if (!cupTotals[cupNum]) {
          cupTotals[cupNum] = { total: 0, count: 0 };
        }
        cupTotals[cupNum].total += total;
        cupTotals[cupNum].count += 1;
      }
    }
    
    var list = [];
    for (var cupNum in cupTotals) {
      var avgTotal = cupTotals[cupNum].total / cupTotals[cupNum].count;
      list.push({
        cup: Number(cupNum),
        avgTotal: avgTotal,
        count: cupTotals[cupNum].count
      });
    }
    
    list.sort(function(a, b) { return a.cup - b.cup; });
    
    return list;
  } catch (e) {
    logError_('getAllCompletedReviews', e);
    return [];
  }
}

// ===== v3.0: 특정 컵의 검수 완료 상세 정보 =====
function getCompletedDetailByCup(cupNum) {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    if (!scoreSheet) return [];
    
    var data = scoreSheet.getDataRange().getValues();
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var mode = String(data[i][5]).toLowerCase().trim();
      var cup = Number(data[i][1]);
      var status = String(data[i][25]).trim();
      
      if (mode === 'judge' && status === '검수완료' && cup === Number(cupNum)) {
        // 역할 정보 가져오기
        var judgeId = String(data[i][2]).trim();
        var role = getUserRole_(judgeId);
        
        results.push({
          judgeName: String(data[i][3]), 
          judgeId: judgeId,
          role: role,
          team: String(data[i][4]),
          flavor: Number(data[i][7]), 
          flavorIntensity: Number(data[i][8]),
          aftertaste: Number(data[i][10]), 
          aftertastePersistence: Number(data[i][11]),
          acidity: Number(data[i][13]), 
          acidityIntensity: Number(data[i][14]),
          body: Number(data[i][16]), 
          bodyIntensity: Number(data[i][17]),
          sweetness: Number(data[i][19]), 
          sweetnessIntensity: Number(data[i][20]),
          overall: Number(data[i][22]),
          notes: { 
            flavor: String(data[i][9]), 
            aftertaste: String(data[i][12]), 
            acidity: String(data[i][15]), 
            body: String(data[i][18]), 
            sweetness: String(data[i][21]), 
            overall: String(data[i][23]) 
          },
          total: Number(data[i][24])
        });
      }
    }
    
    return results;
  } catch (e) {
    logError_('getCompletedDetailByCup', e);
    return [];
  }
}

// Helper: 사용자 역할 조회
function getUserRole_(userId) {
  try {
    var ss = getSS_();
    var userSheet = ss.getSheetByName('Users');
    if (!userSheet) return 'judge';
    
    // userId에서 전화번호 추출 (user_01029090489 → 01029090489)
    var phoneNumber = String(userId).replace('user_', '').trim();
    
    var data = userSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var dbPhone = String(data[i][2]).trim().replace(/[\s-]/g, '');
      
      if (dbPhone === phoneNumber) {
        var role = String(data[i][3]).trim();
        return role; // judge1, judge2, head1 등 그대로 반환
      }
    }
    return 'judge';
  } catch (e) {
    return 'judge';
  }
}

// ===== v3.0: 컵별 모든 평가자 데이터 반환 =====
function getAllCupEvaluations(cupNum) {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    var userSheet = ss.getSheetByName('Users');
    if (!scoreSheet) return [];
    
    // Users 시트에서 role 정보 가져오기
    var userRoles = {};
    if (userSheet) {
      var userData = userSheet.getDataRange().getValues();
      for (var j = 1; j < userData.length; j++) {
        var userId = String(userData[j][0]);
        var role = String(userData[j][3]).trim().toLowerCase();
        // role에서 숫자 제거
        userRoles[userId] = role.replace(/\d+$/, '');
      }
    }
    
    var data = scoreSheet.getDataRange().getValues();
    var evaluations = [];
    
    for (var i = 1; i < data.length; i++) {
      var cup = Number(data[i][1]);
      var mode = String(data[i][5]).toLowerCase();
      var status = String(data[i][25]);
      
      // 해당 컵의 모든 judge 모드 평가 (검수완료된 것만)
      if (cup === Number(cupNum) && mode === 'judge' && status === '검수완료') {
        var judgeId = String(data[i][2]);
        var role = userRoles[judgeId] || 'judge';
        
        evaluations.push({
          judgeName: String(data[i][3]),
          judgeRole: role,
          team: String(data[i][4]),
          process: String(data[i][6]),
          flavor: Number(data[i][7]),
          flavorIntensity: Number(data[i][8]),
          aftertaste: Number(data[i][10]),
          aftertastePersistence: Number(data[i][11]),
          acidity: Number(data[i][13]),
          acidityIntensity: Number(data[i][14]),
          body: Number(data[i][16]),
          bodyIntensity: Number(data[i][17]),
          sweetness: Number(data[i][19]),
          sweetnessIntensity: Number(data[i][20]),
          overall: Number(data[i][22]),
          notes: {
            flavor: String(data[i][9]),
            aftertaste: String(data[i][12]),
            acidity: String(data[i][15]),
            body: String(data[i][18]),
            sweetness: String(data[i][21]),
            overall: String(data[i][23])
          },
          total: Number(data[i][24])
        });
      }
    }
    
    // 총점 내림차순 정렬
    evaluations.sort(function(a, b) { return b.total - a.total; });
    
    return evaluations;
  } catch (e) {
    logError_('getAllCupEvaluations', e);
    return [];
  }
}

// ===== PDF 생성 백엔드 함수 =====

function generateCuppingPDF(cupNumber) {
  try {
    var ss = getSS_();
    var scoreSheet = ss.getSheetByName('Scores');
    
    if (!scoreSheet) {
      return { success: false, message: '평가 데이터가 없습니다.' };
    }
    
    var data = scoreSheet.getDataRange().getValues();
    var evaluations = [];
    
    // 특정 컵 또는 전체 데이터 수집
    for (var i = 1; i < data.length; i++) {
      var cup = Number(data[i][1]);
      var mode = String(data[i][5]).toLowerCase();
      var status = String(data[i][24]);
      
      // Judge 모드 + 검수완료만
      if (mode === 'judge' && status === '검수완료') {
        if (!cupNumber || cup === Number(cupNumber)) {
          evaluations.push({
            cupNumber: cup,
            judgeId: String(data[i][2]),
            judgeName: String(data[i][3]),
            team: String(data[i][4]),
            flavor: Number(data[i][6]),
            flavorIntensity: Number(data[i][7]),
            flavorComment: String(data[i][8]),
            aftertaste: Number(data[i][9]),
            aftertastePersistence: Number(data[i][10]),
            aftertasteComment: String(data[i][11]),
            acidity: Number(data[i][12]),
            acidityIntensity: Number(data[i][13]),
            acidityComment: String(data[i][14]),
            body: Number(data[i][15]),
            bodyIntensity: Number(data[i][16]),
            bodyComment: String(data[i][17]),
            sweetness: Number(data[i][18]),
            sweetnessIntensity: Number(data[i][19]),
            sweetnessComment: String(data[i][20]),
            overall: Number(data[i][21]),
            overallComment: String(data[i][22]),
            total: Number(data[i][23])
          });
        }
      }
    }
    
    if (evaluations.length === 0) {
      return { success: false, message: '출력할 평가 데이터가 없습니다.' };
    }
    
    // 컵별로 그룹화
    var cupGroups = {};
    for (var i = 0; i < evaluations.length; i++) {
      var cup = evaluations[i].cupNumber;
      if (!cupGroups[cup]) cupGroups[cup] = [];
      cupGroups[cup].push(evaluations[i]);
    }
    
    // HTML 생성
    var html = generatePDFHtml(cupGroups);
    
    // HTML을 임시 파일로 저장하고 URL 반환
    var htmlOutput = HtmlService.createHtmlOutput(html)
      .setWidth(800)
      .setHeight(1131); // A4 height in pixels at 96 DPI
    
    return {
      success: true,
      htmlContent: html,
      message: '새 창에서 PDF로 인쇄하세요.'
    };
    
  } catch (e) {
    logError_('generateCuppingPDF', e);
    return { success: false, message: 'PDF 생성 오류: ' + e.toString() };
  }
}

function generatePDFHtml(cupGroups) {
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
  html += '<title>KCR Cupping Report</title>';
  html += '<style>';
  html += '@page { size: A4; margin: 15mm; }';
  html += 'body { font-family: "Noto Sans KR", -apple-system, sans-serif; font-size: 10pt; line-height: 1.4; color: #000; background: #fff; margin: 0; padding: 0; }';
  html += '.page { page-break-after: always; padding: 20px; }';
  html += '.page:last-child { page-break-after: auto; }';
  html += '.cup-title { text-align: center; font-size: 24pt; font-weight: 700; margin-bottom: 10px; border-bottom: 3px solid #000; padding-bottom: 10px; }';
  html += '.stats { background: #f5f5f5; padding: 10px; margin-bottom: 15px; border-radius: 5px; text-align: center; }';
  html += '.stats-row { display: inline-block; margin: 0 15px; font-size: 11pt; }';
  html += '.stats-label { font-weight: 600; }';
  html += '.judges-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 15px; }';
  html += '.judge-card { border: 2px solid #333; border-radius: 8px; padding: 12px; background: #fafafa; page-break-inside: avoid; }';
  html += '.judge-header { font-size: 13pt; font-weight: 700; margin-bottom: 8px; padding-bottom: 6px; border-bottom: 2px solid #666; }';
  html += '.judge-total { float: right; font-size: 16pt; color: #000; }';
  html += '.score-table { width: 100%; border-collapse: collapse; margin-bottom: 8px; font-size: 9pt; }';
  html += '.score-table th { background: #333; color: #fff; padding: 4px; text-align: left; font-weight: 600; font-size: 8pt; }';
  html += '.score-table td { padding: 4px; border-bottom: 1px solid #ddd; }';
  html += '.score-table td:first-child { font-weight: 600; width: 70px; }';
  html += '.score-table td:nth-child(2) { text-align: center; width: 40px; font-weight: 700; }';
  html += '.score-table td:nth-child(3) { text-align: center; width: 35px; font-size: 8pt; color: #666; }';
  html += '.comment-section { margin-top: 6px; }';
  html += '.comment-label { font-size: 8pt; font-weight: 600; color: #666; margin-bottom: 2px; }';
  html += '.comment-text { font-size: 8.5pt; line-height: 1.3; color: #000; background: #fff; padding: 5px; border-left: 3px solid #666; margin-bottom: 4px; }';
  html += '@media print { body { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }';
  html += '</style>';
  html += '</head><body>';
  
  // 각 컵별로 페이지 생성
  for (var cupNum in cupGroups) {
    var judges = cupGroups[cupNum];
    
    // 통계 계산
    var totals = judges.map(function(j) { return j.total; });
    var avgTotal = totals.reduce(function(a, b) { return a + b; }, 0) / totals.length;
    var variance = totals.reduce(function(acc, val) { return acc + Math.pow(val - avgTotal, 2); }, 0) / totals.length;
    var stdDev = Math.sqrt(variance);
    
    html += '<div class="page">';
    html += '<div class="cup-title">Cup ' + cupNum + '</div>';
    html += '<div class="stats">';
    html += '<span class="stats-row"><span class="stats-label">평균:</span> ' + avgTotal.toFixed(2) + '점</span>';
    html += '<span class="stats-row"><span class="stats-label">표준편차:</span> ± ' + stdDev.toFixed(2) + '</span>';
    html += '<span class="stats-row"><span class="stats-label">평가자:</span> ' + judges.length + '명</span>';
    html += '</div>';
    
    html += '<div class="judges-grid">';
    
    // 각 심사위원 카드 (최대 4명)
    for (var i = 0; i < Math.min(judges.length, 4); i++) {
      var j = judges[i];
      
      html += '<div class="judge-card">';
      html += '<div class="judge-header">' + j.judgeName + ' <span class="judge-total">' + j.total.toFixed(1) + '</span></div>';
      
      html += '<table class="score-table">';
      html += '<tr><th>속성</th><th>점수</th><th>강도</th></tr>';
      html += '<tr><td>Flavor</td><td>' + j.flavor.toFixed(2) + '</td><td>(' + j.flavorIntensity + ')</td></tr>';
      html += '<tr><td>Aftertaste</td><td>' + j.aftertaste.toFixed(2) + '</td><td>(' + j.aftertastePersistence + ')</td></tr>';
      html += '<tr><td>Acidity</td><td>' + j.acidity.toFixed(2) + '</td><td>(' + j.acidityIntensity + ')</td></tr>';
      html += '<tr><td>Body</td><td>' + j.body.toFixed(2) + '</td><td>(' + j.bodyIntensity + ')</td></tr>';
      html += '<tr><td>Sweetness</td><td>' + (j.sweetness * 2).toFixed(2) + '</td><td>(' + j.sweetnessIntensity + ')</td></tr>';
      html += '<tr><td>Overall</td><td>' + j.overall.toFixed(2) + '</td><td>-</td></tr>';
      html += '</table>';
      
      html += '<div class="comment-section">';
      if (j.flavorComment) {
        html += '<div class="comment-label">Flavor</div>';
        html += '<div class="comment-text">' + j.flavorComment + '</div>';
      }
      if (j.aftertasteComment) {
        html += '<div class="comment-label">Aftertaste</div>';
        html += '<div class="comment-text">' + j.aftertasteComment + '</div>';
      }
      if (j.acidityComment) {
        html += '<div class="comment-label">Acidity</div>';
        html += '<div class="comment-text">' + j.acidityComment + '</div>';
      }
      if (j.bodyComment) {
        html += '<div class="comment-label">Body</div>';
        html += '<div class="comment-text">' + j.bodyComment + '</div>';
      }
      if (j.sweetnessComment) {
        html += '<div class="comment-label">Sweetness</div>';
        html += '<div class="comment-text">' + j.sweetnessComment + '</div>';
      }
      if (j.overallComment) {
        html += '<div class="comment-label">Overall</div>';
        html += '<div class="comment-text">' + j.overallComment + '</div>';
      }
      html += '</div>';
      
      html += '</div>';
    }
    
    html += '</div>';
    html += '</div>';
  }
  
  html += '</body></html>';
  return html;
}
