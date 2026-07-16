(() => {
    'use strict';

    const STORAGE_PREFIX = 'quizsheet:progress:';
    const STARRED_PREFIX = 'quizsheet:starred:';
    const MAX_FILE_SIZE = 20 * 1024 * 1024;
    const MAX_QUESTION_ROWS = 20000;
    const OPTION_KEYS = 'ABCDEFGH'.split('');
    const HEADER_ALIASES = {
        question: ['题目', '题干', '问题'],
        type: ['题型', '类型'],
        answer: ['答案', '正确答案'],
        explanation: ['解析', '答案解析'],
        category: ['分类', '知识点', '科目'],
        difficulty: ['难度']
    };
    const TRUE_WORDS = ['正确', '对', '是', 'TRUE', 'T', 'YES', 'Y', '√', '✓'];
    const FALSE_WORDS = ['错误', '错', '否', 'FALSE', 'F', 'NO', 'N', '×', '✕', 'X'];

    const elements = {};
    let loadedQuestions = [];
    let pendingImport = null;
    let baseMode = 'sequential';
    let toastTimer = null;
    let advanceTimer = null;
    let importSequence = 0;

    const state = {
        libraryId: '',
        sourceName: '',
        mode: 'sequential',
        questions: [],
        current: 0,
        answers: {},
        statuses: {},
        starred: {},
        revealed: {},
        settingsSignature: '',
        sessionSeed: '',
        autoAdvance: false,
        wrongDelayMs: 3000,
        finished: false
    };

    document.addEventListener('DOMContentLoaded', init);

    function init() {
        [
            'uploadSection', 'uploadTitle', 'fileInput', 'modeSelect', 'loadButton', 'importReport',
            'sessionSettings', 'questionCountSelect', 'categorySelect', 'difficultySelect', 'shuffleOptionsInput', 'autoAdvanceInput', 'wrongDelayInput',
            'quizSection', 'libraryName', 'modeBadge',
            'saveStatus', 'changeFileButton', 'answeredSummary', 'questionGrid', 'progressText',
            'progressFill', 'questionScroll', 'questionContent', 'answerError', 'feedback', 'prevButton',
            'dontKnowButton', 'primaryActionButton', 'resultSection', 'resultTitle', 'scoreValue', 'resultStats',
            'categoryBreakdown', 'resultDetails', 'resultReviewList', 'retryWrongButton',
            'retryStarredButton', 'restartButton', 'resultChangeFileButton', 'toast'
        ].forEach(id => { elements[id] = document.getElementById(id); });

        elements.loadButton.addEventListener('click', onLoadButtonClick);
        elements.fileInput.addEventListener('change', onFileChange);
        elements.prevButton.addEventListener('click', goPrevious);
        elements.dontKnowButton.addEventListener('click', revealAnswer);
        elements.primaryActionButton.addEventListener('click', handlePrimaryAction);
        elements.changeFileButton.addEventListener('click', showUpload);
        elements.resultChangeFileButton.addEventListener('click', showUpload);
        elements.restartButton.addEventListener('click', restartAll);
        elements.retryWrongButton.addEventListener('click', retryWrong);
        elements.retryStarredButton.addEventListener('click', retryStarred);
        [elements.modeSelect, elements.questionCountSelect, elements.categorySelect, elements.difficultySelect]
            .forEach(control => control.addEventListener('change', updatePreparedCount));
        elements.autoAdvanceInput.addEventListener('change', syncAutoAdvanceSettings);
        elements.wrongDelayInput.addEventListener('change', normalizeWrongDelayInput);
        document.addEventListener('pointerdown', () => document.body.classList.remove('keyboard-navigation'), true);
        document.addEventListener('keydown', event => {
            if (!event.ctrlKey && !event.metaKey && !event.altKey) {
                document.body.classList.add('keyboard-navigation');
            }
        }, true);
        document.addEventListener('keydown', handleKeyboard);
        ['pointerdown', 'wheel', 'touchstart'].forEach(eventName => {
            elements.questionScroll.addEventListener(eventName, pauseAutoAdvance, { passive: true });
        });

        if (typeof XLSX === 'undefined') {
            elements.loadButton.disabled = true;
            showImportReport(['Excel 解析组件未能加载，请确认 vendor/xlsx.full.min.js 文件存在。'], 0);
        }
    }

    function onFileChange() {
        importSequence += 1;
        pendingImport = null;
        elements.loadButton.disabled = false;
        elements.loadButton.textContent = '加载题库';
        elements.importReport.classList.add('hidden');
        elements.sessionSettings.classList.add('hidden');
        const file = elements.fileInput.files[0];
        const label = document.querySelector('.file-picker span');
        if (file && label) {
            label.textContent = file.name;
            loadWorkbook();
        }
    }

    function onLoadButtonClick() {
        if (pendingImport) {
            const data = pendingImport;
            beginWithParsedData(data);
            return;
        }
        loadWorkbook();
    }

    function loadWorkbook() {
        const file = elements.fileInput.files[0];
        if (!file) {
            showToast('请先选择一个 .xlsx 文件');
            return;
        }
        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            showImportReport(['仅支持 .xlsx 格式，请将旧版 .xls 文件另存为 .xlsx。'], 0);
            return;
        }
        if (file.size > MAX_FILE_SIZE) {
            showImportReport(['题库文件不能超过 20 MB，请拆分后再导入。'], 0);
            return;
        }

        elements.loadButton.disabled = true;
        elements.loadButton.textContent = '正在解析…';
        const requestId = ++importSequence;
        const reader = new FileReader();
        reader.onload = event => {
            if (requestId !== importSequence) return;
            try {
                const workbook = XLSX.read(new Uint8Array(event.target.result), { type: 'array' });
                if (!workbook.SheetNames.length) throw new Error('文件中没有工作表');
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
                if (rows.length - 1 > MAX_QUESTION_ROWS) {
                    throw new Error(`题目不能超过 ${MAX_QUESTION_ROWS} 行，请拆分题库`);
                }
                const result = parseRows(rows);
                if (!result.questions.length) {
                    showImportReport(result.errors.length ? result.errors : ['没有找到有效题目。'], 0, result.stats);
                    return;
                }

                const parsed = { questions: result.questions, errors: result.errors, file };
                pendingImport = parsed;
                populateSessionSettings(result.questions);
                showImportReport(result.errors, result.questions.length, result.stats);
            } catch (error) {
                console.error(error);
                showImportReport([`解析失败：${error.message || '文件格式异常'}`], 0);
            } finally {
                if (requestId !== importSequence) return;
                elements.loadButton.disabled = false;
                if (!pendingImport) elements.loadButton.textContent = '加载题库';
            }
        };
        reader.onerror = () => {
            if (requestId !== importSequence) return;
            elements.loadButton.disabled = false;
            elements.loadButton.textContent = '加载题库';
            showImportReport(['浏览器读取文件失败，请重新选择文件。'], 0);
        };
        reader.readAsArrayBuffer(file);
    }

    function parseRows(rows) {
        const errors = [];
        const stats = { dataRows: Math.max(rows.length - 1, 0), valid: 0, invalid: 0, duplicates: 0 };
        if (!rows.length) return { questions: [], errors: ['工作表内容为空。'], stats };

        const headers = rows[0].map(value => String(value).trim());
        const columns = {};
        Object.entries(HEADER_ALIASES).forEach(([name, aliases]) => {
            columns[name] = headers.findIndex(header => aliases.includes(header));
        });
        const missing = ['question', 'type', 'answer'].filter(name => columns[name] < 0);
        if (missing.length) {
            const labels = { question: '题目', type: '题型', answer: '答案' };
            stats.invalid = stats.dataRows;
            return { questions: [], errors: [`缺少必需列：${missing.map(name => labels[name]).join('、')}`], stats };
        }

        const optionColumns = OPTION_KEYS
            .map(key => ({ key, index: headers.indexOf(key) }))
            .filter(item => item.index >= 0);
        const questions = [];
        const seenQuestions = new Map();

        rows.slice(1).forEach((row, offset) => {
            const rowNumber = offset + 2;
            if (!row.some(value => String(value).trim() !== '')) return;

            const questionText = cell(row, columns.question);
            const rawType = cell(row, columns.type);
            const rawAnswer = cell(row, columns.answer);
            if (!questionText || !rawType || !rawAnswer) {
                errors.push(`第 ${rowNumber} 行：题目、题型和答案不能为空。`);
                return;
            }

            const type = normalizeType(rawType);
            if (!type) {
                errors.push(`第 ${rowNumber} 行：无法识别题型“${rawType}”。`);
                return;
            }

            let options = optionColumns
                .map(item => ({ key: item.key, value: cell(row, item.index) }))
                .filter(option => option.value !== '');
            if (type === '判断题' && options.length === 0) {
                options = [{ key: 'A', value: '正确' }, { key: 'B', value: '错误' }];
            }
            if (options.length < 2) {
                errors.push(`第 ${rowNumber} 行：至少需要两个有效选项。`);
                return;
            }

            const answerResult = normalizeAnswer(rawAnswer, type, options);
            if (answerResult.error) {
                errors.push(`第 ${rowNumber} 行：${answerResult.error}`);
                return;
            }
            const validKeys = new Set(options.map(option => option.key));
            const invalidKeys = answerResult.keys.filter(key => !validKeys.has(key));
            if (invalidKeys.length) {
                errors.push(`第 ${rowNumber} 行：答案 ${invalidKeys.join('、')} 没有对应选项。`);
                return;
            }
            if (type !== '多选题' && answerResult.keys.length !== 1) {
                errors.push(`第 ${rowNumber} 行：${type}只能有一个正确答案。`);
                return;
            }

            const normalizeDuplicateText = value => String(value).trim().toLocaleLowerCase().replace(/\s+/g, ' ');
            const duplicateKey = JSON.stringify([
                type,
                normalizeDuplicateText(questionText),
                options.map(option => [option.key, normalizeDuplicateText(option.value)]),
                [...answerResult.keys].sort()
            ]);
            if (seenQuestions.has(duplicateKey)) {
                errors.push(`第 ${rowNumber} 行：与第 ${seenQuestions.get(duplicateKey)} 行题目重复，已跳过。`);
                stats.duplicates += 1;
                return;
            }
            seenQuestions.set(duplicateKey, rowNumber);

            questions.push({
                id: `row-${rowNumber}`,
                sourceRow: rowNumber,
                type,
                text: questionText,
                options,
                correctAnswer: answerResult.keys,
                explanation: cell(row, columns.explanation),
                category: cell(row, columns.category),
                difficulty: cell(row, columns.difficulty)
            });
        });
        stats.valid = questions.length;
        stats.invalid = Math.max(errors.length - stats.duplicates, 0);
        return { questions, errors, stats };
    }

    function cell(row, index) {
        return index >= 0 && row[index] !== undefined ? String(row[index]).trim() : '';
    }

    function normalizeType(value) {
        const type = String(value).trim();
        if (type.includes('多选')) return '多选题';
        if (type.includes('单选')) return '单选题';
        if (type.includes('判断')) return '判断题';
        return '';
    }

    function normalizeAnswer(value, type, options) {
        const source = String(value).trim();
        const upper = source.toUpperCase();
        if (type === '判断题' && (TRUE_WORDS.includes(upper) || FALSE_WORDS.includes(upper))) {
            const positive = TRUE_WORDS.includes(upper);
            const wordSet = positive ? TRUE_WORDS : FALSE_WORDS;
            const matchingOption = options.find(option => wordSet.includes(option.value.trim().toUpperCase()));
            return { keys: [matchingOption ? matchingOption.key : (positive ? 'A' : 'B')] };
        }

        const compact = upper.replace(/[，、；;|\s]+/g, ',').replace(/^,+|,+$/g, '');
        if (!compact) return { error: '答案为空或格式不受支持。', keys: [] };
        let tokens = compact.split(',').filter(Boolean);
        if (tokens.length === 1 && /^[A-H]+$/.test(tokens[0])) tokens = tokens[0].split('');
        if (!tokens.length || tokens.some(token => !/^[A-H]$/.test(token))) {
            return { error: `无法识别答案“${source}”，请使用 A、AB 或 A,B 等格式。`, keys: [] };
        }
        return { keys: [...new Set(tokens)] };
    }

    function populateSessionSettings(questions) {
        replaceSelectOptions(elements.questionCountSelect, [
            { value: 'all', label: `全部题目（${questions.length}）` },
            ...[10, 20, 50].filter(count => count < questions.length).map(count => ({ value: String(count), label: `${count} 题` }))
        ]);
        const categories = [...new Set(questions.map(question => question.category).filter(Boolean))].sort();
        const difficulties = [...new Set(questions.map(question => question.difficulty).filter(Boolean))].sort();
        replaceSelectOptions(elements.categorySelect, [
            { value: 'all', label: '全部分类' },
            ...categories.map(value => ({ value, label: value }))
        ]);
        replaceSelectOptions(elements.difficultySelect, [
            { value: 'all', label: '全部难度' },
            ...difficulties.map(value => ({ value, label: value }))
        ]);
        elements.sessionSettings.classList.remove('hidden');
        syncAutoAdvanceSettings();
        updatePreparedCount();
    }

    function replaceSelectOptions(select, items) {
        select.replaceChildren(...items.map(item => {
            const option = document.createElement('option');
            option.value = item.value;
            option.textContent = item.label;
            return option;
        }));
    }

    function currentSessionSettings() {
        const wrongDelaySeconds = normalizeWrongDelayInput();
        return {
            mode: elements.modeSelect.value,
            count: elements.questionCountSelect.value,
            category: elements.categorySelect.value,
            difficulty: elements.difficultySelect.value,
            shuffleOptions: elements.shuffleOptionsInput.checked,
            autoAdvance: elements.autoAdvanceInput.checked,
            wrongDelaySeconds
        };
    }

    function syncAutoAdvanceSettings() {
        elements.wrongDelayInput.disabled = !elements.autoAdvanceInput.checked;
    }

    function normalizeWrongDelayInput() {
        const parsed = Number(elements.wrongDelayInput.value);
        const clamped = Number.isFinite(parsed) ? Math.min(5, Math.max(0, parsed)) : 3;
        const seconds = Math.round(clamped * 2) / 2;
        elements.wrongDelayInput.value = String(seconds);
        return seconds;
    }

    function updatePreparedCount() {
        if (!pendingImport) return;
        const settings = currentSessionSettings();
        const available = pendingImport.questions.filter(question =>
            (settings.category === 'all' || question.category === settings.category) &&
            (settings.difficulty === 'all' || question.difficulty === settings.difficulty)
        ).length;
        const count = settings.count === 'all' ? available : Math.min(Number(settings.count), available);
        elements.loadButton.textContent = count ? `开始练习（${count} 题）` : '当前筛选无题目';
        elements.loadButton.disabled = count === 0;
    }

    function prepareQuestions(questions, settings, libraryId, sessionSeed) {
        let selected = questions.filter(question =>
            (settings.category === 'all' || question.category === settings.category) &&
            (settings.difficulty === 'all' || question.difficulty === settings.difficulty)
        );
        if (settings.mode === 'random') {
            seededShuffle(selected, `${libraryId}:${sessionSeed}:questions:${settings.count}:${settings.category}:${settings.difficulty}`);
        }
        if (settings.count !== 'all') selected = selected.slice(0, Number(settings.count));
        return selected.map(question => settings.shuffleOptions && question.type !== '判断题'
            ? shuffleQuestionOptions(question, `${libraryId}:${sessionSeed}:${question.id}`)
            : { ...question, options: question.options.map(option => ({ ...option })), correctAnswer: [...question.correctAnswer] });
    }

    function shuffleQuestionOptions(question, seedText) {
        const options = question.options.map(option => ({ ...option }));
        let seed = parseInt(fingerprint(seedText), 16) || 1;
        for (let index = options.length - 1; index > 0; index -= 1) {
            seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
            const target = seed % (index + 1);
            [options[index], options[target]] = [options[target], options[index]];
        }
        const remappedOptions = options.map((option, index) => ({ key: OPTION_KEYS[index], value: option.value, originalKey: option.key }));
        const correctAnswer = remappedOptions
            .filter(option => question.correctAnswer.includes(option.originalKey))
            .map(option => option.key);
        return {
            ...question,
            options: remappedOptions.map(({ key, value }) => ({ key, value })),
            correctAnswer
        };
    }

    function seededShuffle(items, seedText) {
        let seed = parseInt(fingerprint(seedText), 16) || 1;
        for (let index = items.length - 1; index > 0; index -= 1) {
            seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0;
            const target = seed % (index + 1);
            [items[index], items[target]] = [items[target], items[index]];
        }
    }

    function beginWithParsedData({ questions, errors, file }) {
        state.libraryId = fingerprint(JSON.stringify(questions));
        state.sourceName = file.name;
        const settings = currentSessionSettings();
        baseMode = settings.mode;
        state.autoAdvance = settings.autoAdvance;
        state.wrongDelayMs = Math.round(settings.wrongDelaySeconds * 1000);
        state.settingsSignature = fingerprint(JSON.stringify(settings));
        const saved = readProgressRecord();
        state.starred = restoreStarred();
        if (saved?.starred && typeof saved.starred === 'object' && !Array.isArray(saved.starred)) {
            state.starred = { ...saved.starred, ...state.starred };
            persistStarred();
        }
        state.sessionSeed = saved?.settingsSignature === state.settingsSignature && saved.sessionSeed
            ? saved.sessionSeed
            : createSessionSeed();
        loadedQuestions = prepareQuestions(questions, settings, state.libraryId, state.sessionSeed);
        if (!loadedQuestions.length) {
            pendingImport = { questions, errors, file };
            showToast('当前筛选条件下没有题目，请调整练习设置');
            return false;
        }
        state.revealed = {};

        const restored = restoreProgress();
        if (!restored) startFresh(baseMode);
        showQuiz();
        if (restored) showToast('已恢复这份题库的上次进度');
        else if (errors.length) showToast(`已跳过 ${errors.length} 行无效数据`);
        return true;
    }

    function startFresh(mode, subset = loadedQuestions) {
        cancelAutoAdvance();
        state.mode = mode;
        state.questions = subset.slice();
        if (mode === 'random') shuffle(state.questions);
        state.current = 0;
        state.answers = {};
        state.statuses = {};
        state.revealed = {};
        state.finished = false;
        persistProgress();
    }

    function restoreProgress() {
        try {
            const saved = readProgressRecord();
            if (!saved || !Array.isArray(saved.orderIds) || !saved.orderIds.length) return false;
            if (saved.settingsSignature !== state.settingsSignature) return false;
            const byId = new Map(loadedQuestions.map(question => [question.id, question]));
            const ordered = saved.orderIds.map(id => byId.get(id)).filter(Boolean);
            if (!ordered.length || ordered.length !== saved.orderIds.length) return false;
            state.mode = ['sequential', 'random', 'wrong', 'starred'].includes(saved.mode) ? saved.mode : baseMode;
            baseMode = ['sequential', 'random'].includes(saved.baseMode)
                ? saved.baseMode
                : (state.mode === 'random' ? 'random' : 'sequential');
            state.questions = ordered;
            state.current = Math.min(Math.max(Number(saved.current) || 0, 0), ordered.length - 1);
            state.answers = saved.answers && typeof saved.answers === 'object' ? saved.answers : {};
            state.statuses = saved.statuses && typeof saved.statuses === 'object' ? saved.statuses : {};
            state.revealed = saved.revealed && typeof saved.revealed === 'object' ? saved.revealed : {};
            state.finished = Boolean(saved.finished) && ordered.every(question => isSubmitted(question));
            if (saved.finished && !state.finished) {
                state.current = ordered.findIndex(question => !isSubmitted(question));
            }
            return true;
        } catch (error) {
            console.warn('无法恢复进度', error);
            return false;
        }
    }

    function persistProgress() {
        if (!state.libraryId || !state.questions.length) return;
        try {
            localStorage.setItem(storageKey(), JSON.stringify({
                mode: state.mode,
                baseMode,
                settingsSignature: state.settingsSignature,
                sessionSeed: state.sessionSeed,
                autoAdvance: state.autoAdvance,
                wrongDelayMs: state.wrongDelayMs,
                orderIds: state.questions.map(question => question.id),
                current: state.current,
                answers: state.answers,
                statuses: state.statuses,
                revealed: state.revealed,
                finished: state.finished
            }));
            if (elements.saveStatus) elements.saveStatus.textContent = '已自动保存';
        } catch (error) {
            console.warn('无法保存进度', error);
            if (elements.saveStatus) elements.saveStatus.textContent = '进度保存失败';
        }
    }

    function storageKey() {
        return `${STORAGE_PREFIX}${state.libraryId}`;
    }

    function readProgressRecord() {
        try {
            return JSON.parse(localStorage.getItem(storageKey()) || 'null');
        } catch {
            return null;
        }
    }

    function starredStorageKey() {
        return `${STARRED_PREFIX}${state.libraryId}`;
    }

    function restoreStarred() {
        try {
            const stored = JSON.parse(localStorage.getItem(starredStorageKey()) || '{}');
            return stored && typeof stored === 'object' && !Array.isArray(stored) ? stored : {};
        } catch {
            return {};
        }
    }

    function persistStarred() {
        if (!state.libraryId) return;
        try {
            localStorage.setItem(starredStorageKey(), JSON.stringify(state.starred));
        } catch (error) {
            console.warn('无法保存存疑标记', error);
        }
    }

    function createSessionSeed() {
        return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
    }

    function preferredScrollBehavior() {
        return window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth';
    }

    function focusForKeyboard(element) {
        if (element && document.body.classList.contains('keyboard-navigation')) {
            window.requestAnimationFrame(() => element.focus());
        }
    }

    function focusQuestionTitle() {
        focusForKeyboard(document.getElementById('questionTitle'));
    }

    function fingerprint(text) {
        let hash = 2166136261;
        for (let index = 0; index < text.length; index += 1) {
            hash ^= text.charCodeAt(index);
            hash = Math.imul(hash, 16777619);
        }
        return (hash >>> 0).toString(16);
    }

    function shuffle(items) {
        for (let index = items.length - 1; index > 0; index -= 1) {
            const target = Math.floor(Math.random() * (index + 1));
            [items[index], items[target]] = [items[target], items[index]];
        }
    }

    function showQuiz() {
        document.body.classList.add('quiz-active');
        elements.uploadSection.classList.add('hidden');
        elements.resultSection.classList.add('hidden');
        elements.quizSection.classList.remove('hidden');
        elements.libraryName.textContent = state.sourceName;
        elements.modeBadge.textContent = `${modeLabel()}${state.autoAdvance ? ' · 自动跳题' : ''}`;
        if (state.finished) renderResult();
        else {
            renderAll();
            focusQuestionTitle();
        }
    }

    function showUpload() {
        cancelAutoAdvance();
        persistProgress();
        importSequence += 1;
        pendingImport = null;
        elements.fileInput.value = '';
        elements.loadButton.disabled = false;
        elements.loadButton.textContent = '加载题库';
        elements.importReport.classList.add('hidden');
        elements.sessionSettings.classList.add('hidden');
        const fileLabel = document.querySelector('.file-picker span');
        if (fileLabel) fileLabel.textContent = '选择 Excel 题库';
        document.body.classList.remove('quiz-active');
        elements.quizSection.classList.add('hidden');
        elements.resultSection.classList.add('hidden');
        elements.uploadSection.classList.remove('hidden');
        focusForKeyboard(elements.uploadTitle);
    }

    function modeLabel() {
        if (state.mode === 'random') return '随机题序';
        if (state.mode === 'wrong') return '错题复习';
        if (state.mode === 'starred') return '存疑复习';
        return '顺序练习';
    }

    function renderAll() {
        renderQuestionGrid();
        renderQuestion();
        renderProgress();
    }

    function renderQuestionGrid() {
        elements.questionGrid.replaceChildren();
        state.questions.forEach((question, index) => {
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'question-number';
            button.textContent = String(index + 1);
            button.setAttribute('aria-label', `第 ${index + 1} 题，${statusLabel(state.statuses[question.id])}`);
            if (index === state.current) button.classList.add('current');
            if (index === state.current) button.setAttribute('aria-current', 'step');
            if (state.statuses[question.id]) button.classList.add(state.statuses[question.id]);
            button.addEventListener('click', () => navigateTo(index, { requireAnswer: index > state.current }));
            elements.questionGrid.appendChild(button);
        });
        const answered = state.questions.filter(question => state.statuses[question.id]).length;
        elements.answeredSummary.textContent = `${answered} / ${state.questions.length}`;
    }

    function statusLabel(status) {
        if (status === 'correct') return '回答正确';
        if (status === 'incorrect') return '回答错误';
        return '未作答';
    }

    function renderQuestion() {
        const question = state.questions[state.current];
        const submitted = isSubmitted(question);
        const selected = state.answers[question.id] || [];
        elements.questionContent.replaceChildren();

        const meta = document.createElement('div');
        meta.className = 'question-meta';
        [question.type, question.category, question.difficulty].filter(Boolean).forEach(text => {
            const tag = document.createElement('span');
            tag.textContent = text;
            meta.appendChild(tag);
        });
        const starButton = document.createElement('button');
        starButton.type = 'button';
        starButton.className = 'star-button';
        starButton.textContent = state.starred[question.id] ? '★ 已标记存疑' : '☆ 标记存疑';
        starButton.setAttribute('aria-pressed', state.starred[question.id] ? 'true' : 'false');
        starButton.addEventListener('click', toggleStarred);
        meta.appendChild(starButton);

        const title = document.createElement('h2');
        title.id = 'questionTitle';
        title.tabIndex = -1;
        title.className = 'question-title';
        title.textContent = question.text;

        const options = document.createElement('div');
        options.className = 'options';
        if (submitted) options.classList.add('submitted');
        options.setAttribute('role', 'group');
        options.setAttribute('aria-labelledby', 'questionTitle');
        options.setAttribute('aria-describedby', 'answerError');
        question.options.forEach(option => {
            const label = document.createElement('label');
            label.className = 'option-label';
            const isSelected = selected.includes(option.key);
            if (isSelected) label.classList.add('selected');
            if (submitted && question.correctAnswer.includes(option.key)) label.classList.add('correct-answer');
            if (submitted && isSelected && !question.correctAnswer.includes(option.key)) label.classList.add('wrong-answer');

            const input = document.createElement('input');
            input.type = question.type === '多选题' ? 'checkbox' : 'radio';
            input.name = `question-${question.id}`;
            input.value = option.key;
            input.checked = isSelected;
            input.disabled = submitted;
            input.addEventListener('change', captureSelection);

            const key = document.createElement('span');
            key.className = 'option-key';
            key.textContent = `${option.key}.`;
            const value = document.createElement('span');
            value.textContent = option.value;
            label.append(input, key, value);
            options.appendChild(label);
        });
        elements.questionContent.append(meta, title, options);

        elements.answerError.classList.add('hidden');
        renderFeedback(question, submitted);
        elements.prevButton.disabled = state.current === 0;
        elements.dontKnowButton.classList.toggle('hidden', submitted);
        elements.primaryActionButton.textContent = submitted
            ? (state.current === state.questions.length - 1 ? '完成练习' : '下一题')
            : '提交答案';
    }

    function captureSelection() {
        const question = state.questions[state.current];
        const inputs = elements.questionContent.querySelectorAll('input:checked');
        state.answers[question.id] = Array.from(inputs, input => input.value);
        elements.questionContent.querySelectorAll('.option-label').forEach(label => {
            label.classList.toggle('selected', label.querySelector('input').checked);
        });
        elements.answerError.classList.add('hidden');
        persistProgress();
        if (state.autoAdvance && question.type !== '多选题') {
            submitCurrentAnswer({ allowAutoAdvance: true });
        }
    }

    function toggleStarred() {
        const question = state.questions[state.current];
        if (state.starred[question.id]) delete state.starred[question.id];
        else state.starred[question.id] = true;
        persistStarred();
        persistProgress();
        renderQuestion();
        const starButton = elements.questionContent.querySelector('.star-button');
        if (starButton && document.body.classList.contains('keyboard-navigation')) starButton.focus();
    }

    function isSubmitted(question) {
        return ['correct', 'incorrect'].includes(state.statuses[question.id]);
    }

    function promptForAnswer() {
        elements.answerError.textContent = '请先选择至少一个答案';
        elements.answerError.classList.remove('hidden');
        const firstOption = elements.questionContent.querySelector('input:not(:disabled)');
        if (firstOption) firstOption.focus();
    }

    function submitCurrentAnswer({ allowAutoAdvance = false } = {}) {
        const question = state.questions[state.current];
        if (isSubmitted(question)) return true;
        const selected = state.answers[question.id] || [];
        if (!selected.length) {
            promptForAnswer();
            return false;
        }
        const actual = [...selected].sort();
        const expected = [...question.correctAnswer].sort();
        state.statuses[question.id] = actual.length === expected.length && actual.every((key, index) => key === expected[index])
            ? 'correct'
            : 'incorrect';
        persistProgress();
        renderAll();
        if (document.body.classList.contains('keyboard-navigation')) elements.primaryActionButton.focus();
        window.requestAnimationFrame(() => {
            elements.feedback.scrollIntoView({ block: 'nearest', behavior: preferredScrollBehavior() });
        });
        if (allowAutoAdvance && state.autoAdvance) scheduleAutoAdvance(question);
        return true;
    }

    function revealAnswer() {
        const question = state.questions[state.current];
        if (isSubmitted(question)) return;
        state.answers[question.id] = [];
        state.statuses[question.id] = 'incorrect';
        state.revealed[question.id] = true;
        persistProgress();
        renderAll();
        if (document.body.classList.contains('keyboard-navigation')) elements.primaryActionButton.focus();
        window.requestAnimationFrame(() => {
            elements.feedback.scrollIntoView({ block: 'nearest', behavior: preferredScrollBehavior() });
        });
    }

    function renderFeedback(question, submitted) {
        elements.feedback.replaceChildren();
        elements.feedback.className = 'feedback';
        if (!submitted) {
            elements.feedback.classList.add('hidden');
            return;
        }
        const correct = state.statuses[question.id] === 'correct';
        elements.feedback.classList.add(correct ? 'correct' : 'incorrect');
        const result = document.createElement('strong');
        result.textContent = correct
            ? '回答正确'
            : `${state.revealed[question.id] ? '已查看答案' : '回答错误'}，正确答案：${question.correctAnswer.join('、')}`;
        elements.feedback.appendChild(result);
        if (question.explanation) {
            const explanation = document.createElement('div');
            explanation.className = 'explanation';
            explanation.textContent = `解析：${question.explanation}`;
            elements.feedback.appendChild(explanation);
        }
    }

    function renderProgress() {
        const answered = state.questions.filter(question => state.statuses[question.id]).length;
        const percent = state.questions.length ? Math.round(answered / state.questions.length * 100) : 0;
        elements.progressText.textContent = `第 ${state.current + 1} 题，共 ${state.questions.length} 题 · 已答 ${answered} 题`;
        elements.progressFill.style.width = `${percent}%`;
        const progress = elements.progressFill.parentElement;
        progress.setAttribute('aria-valuenow', String(percent));
    }

    function moveTo(index) {
        if (index < 0 || index >= state.questions.length) return;
        cancelAutoAdvance();
        state.current = index;
        persistProgress();
        renderAll();
        elements.questionScroll.scrollTop = 0;
        window.scrollTo({ top: 0, behavior: preferredScrollBehavior() });
        focusQuestionTitle();
    }

    function navigateTo(index, { requireAnswer = false } = {}) {
        if (index < 0 || index >= state.questions.length || index === state.current) return;
        const question = state.questions[state.current];
        if (!isSubmitted(question)) {
            const selected = state.answers[question.id] || [];
            if (requireAnswer) {
                if (selected.length) submitCurrentAnswer();
                else promptForAnswer();
                return;
            }
        }
        moveTo(index);
    }

    function goPrevious() {
        navigateTo(state.current - 1);
    }

    function handlePrimaryAction() {
        const question = state.questions[state.current];
        if (!isSubmitted(question)) {
            submitCurrentAnswer({ allowAutoAdvance: true });
            return;
        }
        if (state.current < state.questions.length - 1) moveTo(state.current + 1);
        else finishQuiz();
    }

    function finishQuiz() {
        cancelAutoAdvance();
        const firstUnanswered = state.questions.findIndex(question => !isSubmitted(question));
        if (firstUnanswered >= 0) {
            state.finished = false;
            moveTo(firstUnanswered);
            const remaining = state.questions.filter(question => !isSubmitted(question)).length;
            showToast(`还有 ${remaining} 道题未作答，已定位到第一道`);
            return;
        }
        state.finished = true;
        persistProgress();
        renderResult();
    }

    function scheduleAutoAdvance(question) {
        cancelAutoAdvance();
        const questionId = question.id;
        const questionIndex = state.current;
        const delay = state.statuses[questionId] === 'correct' ? 450 : state.wrongDelayMs;
        advanceTimer = window.setTimeout(() => {
            advanceTimer = null;
            if (state.current !== questionIndex || state.questions[state.current]?.id !== questionId) return;
            if (state.current < state.questions.length - 1) moveTo(state.current + 1);
            else finishQuiz();
        }, delay);
    }

    function cancelAutoAdvance() {
        if (advanceTimer !== null) {
            window.clearTimeout(advanceTimer);
            advanceTimer = null;
        }
    }

    function pauseAutoAdvance() {
        if (advanceTimer === null) return;
        cancelAutoAdvance();
        showToast('已暂停自动跳题，可手动进入下一题');
    }

    function renderResult() {
        document.body.classList.remove('quiz-active');
        elements.quizSection.classList.add('hidden');
        elements.uploadSection.classList.add('hidden');
        elements.resultSection.classList.remove('hidden');
        const total = state.questions.length;
        const correct = state.questions.filter(question => state.statuses[question.id] === 'correct').length;
        const incorrect = state.questions.filter(question => state.statuses[question.id] === 'incorrect').length;
        const starred = state.questions.filter(question => state.starred[question.id]).length;
        const allStarred = loadedQuestions.filter(question => state.starred[question.id]).length;
        const score = total ? Math.round(correct / total * 100) : 0;
        elements.scoreValue.textContent = String(score);
        elements.resultStats.replaceChildren(
            createStat(correct, '答对'),
            createStat(incorrect, '答错'),
            createStat(starred, '存疑')
        );
        elements.retryWrongButton.classList.toggle('hidden', incorrect === 0);
        elements.retryWrongButton.textContent = `重练 ${incorrect} 道错题`;
        elements.retryStarredButton.classList.toggle('hidden', allStarred === 0);
        elements.retryStarredButton.textContent = `复习 ${allStarred} 道存疑题`;
        renderCategoryBreakdown();
        renderResultDetails();
        focusForKeyboard(elements.resultTitle);
    }

    function createStat(value, label) {
        const item = document.createElement('div');
        item.className = 'stat';
        const number = document.createElement('strong');
        number.textContent = String(value);
        const caption = document.createElement('span');
        caption.textContent = label;
        item.append(number, caption);
        return item;
    }

    function renderCategoryBreakdown() {
        const groups = new Map();
        state.questions.forEach(question => {
            if (!question.category) return;
            if (!groups.has(question.category)) groups.set(question.category, { total: 0, correct: 0 });
            const group = groups.get(question.category);
            group.total += 1;
            if (state.statuses[question.id] === 'correct') group.correct += 1;
        });
        elements.categoryBreakdown.replaceChildren();
        elements.categoryBreakdown.classList.toggle('hidden', groups.size < 2);
        if (groups.size < 2) return;
        const title = document.createElement('h2');
        title.textContent = '分类表现';
        elements.categoryBreakdown.appendChild(title);
        groups.forEach((group, category) => {
            const row = document.createElement('div');
            row.className = 'category-row';
            const label = document.createElement('span');
            label.textContent = category;
            const value = document.createElement('strong');
            value.textContent = `${group.correct} / ${group.total}`;
            const track = document.createElement('div');
            track.className = 'mini-progress';
            const fill = document.createElement('i');
            fill.style.width = `${Math.round(group.correct / group.total * 100)}%`;
            track.appendChild(fill);
            row.append(label, track, value);
            elements.categoryBreakdown.appendChild(row);
        });
    }

    function renderResultDetails() {
        const wrongQuestions = state.questions.filter(question => state.statuses[question.id] === 'incorrect');
        elements.resultDetails.classList.toggle('hidden', wrongQuestions.length === 0);
        elements.resultDetails.open = false;
        elements.resultDetails.querySelector('summary').textContent = `查看 ${wrongQuestions.length} 道错题详情`;
        elements.resultReviewList.replaceChildren(...wrongQuestions.map((question, index) => createReviewCard(question, index)));
    }

    function createReviewCard(question, index) {
        const card = document.createElement('article');
        card.className = 'review-card';
        const title = document.createElement('h3');
        title.textContent = `${index + 1}. ${question.text}`;
        const userAnswer = document.createElement('p');
        userAnswer.textContent = `你的答案：${state.revealed[question.id] ? '不会 / 查看答案' : formatAnswer(question, state.answers[question.id])}`;
        const correctAnswer = document.createElement('p');
        correctAnswer.textContent = `正确答案：${formatAnswer(question, question.correctAnswer)}`;
        card.append(title, userAnswer, correctAnswer);
        if (question.explanation) {
            const explanation = document.createElement('p');
            explanation.className = 'review-explanation';
            explanation.textContent = `解析：${question.explanation}`;
            card.appendChild(explanation);
        }
        return card;
    }

    function formatAnswer(question, keys = []) {
        if (!Array.isArray(keys) || !keys.length) return '未选择';
        return keys.map(key => {
            const option = question.options.find(item => item.key === key);
            return option ? `${key}. ${option.value}` : key;
        }).join('；');
    }

    function retryWrong() {
        const wrong = state.questions.filter(question => state.statuses[question.id] === 'incorrect');
        if (!wrong.length) return;
        startFresh('wrong', wrong);
        showQuiz();
    }

    function retryStarred() {
        const starredQuestions = loadedQuestions.filter(question => state.starred[question.id]);
        if (!starredQuestions.length) return;
        startFresh('starred', starredQuestions);
        showQuiz();
    }

    function restartAll() {
        state.sessionSeed = createSessionSeed();
        if (pendingImport) {
            loadedQuestions = prepareQuestions(pendingImport.questions, currentSessionSettings(), state.libraryId, state.sessionSeed);
        }
        startFresh(baseMode);
        showQuiz();
    }

    function showImportReport(errors, validCount, stats = {}) {
        elements.importReport.replaceChildren();
        elements.importReport.className = `import-report${errors.length ? '' : ' success'}`;
        const title = document.createElement('strong');
        title.textContent = validCount
            ? `解析完成：有效 ${validCount} 题，格式错误 ${stats.invalid || 0} 行，重复 ${stats.duplicates || 0} 题。`
            : '题库未能导入：';
        elements.importReport.appendChild(title);
        if (!errors.length) return;
        const list = document.createElement('ul');
        errors.slice(0, 30).forEach(message => {
            const item = document.createElement('li');
            item.textContent = message;
            list.appendChild(item);
        });
        if (errors.length > 30) {
            const item = document.createElement('li');
            item.textContent = `其余 ${errors.length - 30} 条已省略，请继续检查题库。`;
            list.appendChild(item);
        }
        elements.importReport.appendChild(list);
    }

    function handleKeyboard(event) {
        if (elements.quizSection.classList.contains('hidden') || event.ctrlKey || event.metaKey || event.altKey) return;
        const targetTag = event.target.tagName;
        if (['SELECT', 'TEXTAREA'].includes(targetTag)) return;
        const isOptionControl = targetTag === 'INPUT';
        const isButton = targetTag === 'BUTTON';
        const question = state.questions[state.current];
        const key = event.key.toUpperCase();

        if (/^[A-H]$/.test(key) && !isSubmitted(question)) {
            const input = elements.questionContent.querySelector(`input[value="${key}"]`);
            if (input) {
                event.preventDefault();
                input.click();
            }
            return;
        }
        if (event.key === 'Enter' && !isButton) {
            event.preventDefault();
            handlePrimaryAction();
        } else if (event.key === 'ArrowLeft' && !isOptionControl && !isButton) {
            event.preventDefault();
            goPrevious();
        } else if (event.key === 'ArrowRight' && !isOptionControl && !isButton && isSubmitted(question)) {
            event.preventDefault();
            handlePrimaryAction();
        } else if (event.key === 'Escape' && !isSubmitted(question)) {
            event.preventDefault();
            state.answers[question.id] = [];
            persistProgress();
            renderQuestion();
        }
    }

    function showToast(message) {
        window.clearTimeout(toastTimer);
        elements.toast.textContent = message;
        elements.toast.classList.remove('hidden');
        toastTimer = window.setTimeout(() => elements.toast.classList.add('hidden'), 2600);
    }
})();
