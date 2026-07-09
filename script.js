// ============================================================
// ГИБРИДНЫЙ ДАШБОРД - С ДОБАВЛЕНИЕМ НКИТК И РАСШИРЕННОЙ ФИЛЬТРАЦИЕЙ ДАТ
// ============================================================

// ============================================================
// КОНСТАНТЫ И КОНФИГУРАЦИЯ
// ============================================================

const EXCEL_FILE_PATH = window.EXCEL_FILE_PATH_OVERRIDE || '/file/Dashboard_report_Technical_unit.xlsm';

// Переключатель отладочных логов. Для отладки из консоли: window.DEBUG = true; location.reload();
const DEBUG = window.DEBUG === true;
const log = DEBUG ? console.log.bind(console) : () => {};

// Текущее состояние
let workbook = null;
let worksheetCache = null;
let currentSheet = 'ЦНГО';
let fileLastModified = null;
let isNKiTKMode = false;
let techBlockStatusRulesCache = null;
let workbookDataCache = createWorkbookDataCache();
let productionProgramFilters = {
    businessDirection: '',
    projectStatus: '',
    costCenter: '',
    designBureau: ''
};
let productionProgramFactVisible = false;
let productionProgramDetailsExpanded = false;
let projectsTargetIndicatorOnly = false;
let projectsGroupedByTask = false;
let targetIndicatorPeriodKey = '2026-H1';
let projectsOverdueFilter = 'all';
let expandedProjectDetailsKey = null;
const QUARTERLY_TREND_DISPLAY_LIMIT = 4;
const TODAY_ABSENCE_STATUS_MAP = {
    'к': 'businessTrip',
    'отп': 'vacation',
    'б': 'sickLeave',
    'обс': 'unpaidVacation',
    'ог': 'dayOff',
    'д': 'maternityLeave',
    'р': 'maternityLeave',
    'декрет': 'maternityLeave'
};
const TODAY_ABSENCE_LABELS = {
    businessTrip: 'Командировка',
    vacation: 'Отпуск',
    sickLeave: 'Больничный',
    unpaidVacation: 'Отпуск без сохранения',
    maternityLeave: 'Декрет',
    dayOff: 'Отгул'
};
const TODAY_ABSENCE_ORDER = [
    'businessTrip',
    'vacation',
    'sickLeave',
    'unpaidVacation',
    'maternityLeave',
    'dayOff'
];

function createWorkbookDataCache() {
    return {
        worksheetRows: new Map(),
        worksheetCellCaches: new Map(),
        employeeLookup: null,
        loadByProject: null,
        timesheetCalendar: null,
        projectPlanFact: null,
        projectNamesLookup: null,
        centerCapacityByRange: new Map(),
        productionProgramHtml: new Map()
    };
}

function resetWorkbookDataCache() {
    workbookDataCache = createWorkbookDataCache();
}

// Данные для фильтрации НКиТК
let nkitkTableData = [];
let nkitkFilters = {};               // { [colIndex]: { type, checkedValues?, dateFilter?, customRange? } }
let nkitkSort = null;                // { col, order } — единая сортировка (отдельно от фильтров)
let nkitkGlobalSearch = '';          // строка глобального поиска по всем столбцам
let currentFilterCol = null;
let globalSearchDebounceId = null;

// Менеджер графиков
let chartManager = {
    charts: {},
    
    destroy(id) {
        if (this.charts[id]) {
            try {
                this.charts[id].destroy();
            } catch (e) {
                console.warn('Ошибка при уничтожении графика:', e);
            }
            delete this.charts[id];
        }
    },
    
    destroyAll() {
        Object.values(this.charts).forEach(chart => {
            try {
                chart.destroy();
            } catch (e) {
                console.warn('Ошибка при уничтожении графика:', e);
            }
        });
        this.charts = {};
    },
    
    set(id, chart) {
        this.destroy(id);
        this.charts[id] = chart;
        return chart;
    }
};

// ============================================================
// КОНФИГУРАЦИЯ ЯЧЕК
// ============================================================

const EXCEL_CONFIG = {
    // Заголовки графиков (подписи над диаграммами)
    chartTitles: [
        { id: 'last-week-chart-title', cell: 'A20' },    // Ячейка с заголовком для графика "Нарастающий итог по дням за месяц"
        { id: 'week-chart-title', cell: 'A26' },         // Ячейка с заголовком для графика "Текущий месяц" 
        { id: 'last-month-chart-title', cell: 'A32' },   // Ячейка с заголовком для графика "Прошлый месяц"
        { id: 'month-chart-title', cell: 'A38' },        // Ячейка с заголовком для графика "Текущий месяц" (дубль)
        { id: 'next-month-chart-title', cell: 'A45' }    // Ячейка с заголовком для графика "Следующий месяц"
    ],
    
    // Шапка дашборда - основная информация о центре
    header: {
        centerName: 'A2',    // Название центра (например: "ЦНГО", "ТехБлок")
        manager: 'C2',       // ФИО руководителя центра
        admin: 'D2',         // ФИО администратора центра
        employeeCount: 'M6', // Общее количество сотрудников в центре
        workloadTotal: 'N6'  // Общая загрузка центра (человеко-часы)
    },
    
    // Уведомления и контрольные показатели по проектам
    notification: {
        rows: [
            { key: 'deadline', labelCell: 'BM15', valueCell: 'BN15', valueType: 'projects' },      // Проекты с ближайшими сроками (количество)
            { key: 'hoursProjects', labelCell: 'BM16', valueCell: 'BN16', valueType: 'projects' }, // Проекты с превышением часов (количество)
            { key: 'hoursTotal', labelCell: 'BM17', valueCell: 'BN17', valueType: 'hours' },       // Превышение часов в часах (суммарно)
            { key: 'missingPlan', labelCell: 'BM18', valueCell: 'BN18', valueType: 'projects' }    // Проекты без плановых часов (количество)
        ],
        fallbackCell: 'BN19'  // Резервная ячейка для отображения текста, если нет других уведомлений
    },
    
    // Данные об отсутствиях сотрудников
    absences: {
        names: { column: 'B', startRow: 10, endRow: 15 },     // Столбец B, строки 10-15: названия типов отсутствий (отпуск, больничный и т.д.)
        counts: { column: 'N', startRow: 10, endRow: 15 },    // Столбец N, строки 10-15: количество отсутствий по каждому типу
        employees: { startCol: 'Y', endCol: 'AD', startRow: 6 } // Диапазон Y6:AD15: список сотрудников по типам отсутствий
    },
    
    // Данные о загрузке сотрудников (требуют внимания)
    loading: {
        loadedCell: 'N17',           // Ячейка N17: общее количество загруженных сотрудников
        notLoadedCell: 'N18',        // Ячейка N18: количество сотрудников, требующих внимания (не загружены)
        employeesColumn: 'AE',       // Столбец AE: список сотрудников, требующих внимания
        employeesStartRow: 6         // Начальная строка для списка сотрудников (AE6 и далее)
    },
    
    // Конфигурация графиков загрузки по периодам
    charts: {
        periods: [
            {
                id: 'day',                   // Период
                loadCell: 'N20',             // Фактические часы
                totalCell: 'N21',            // Плановые часы
                idleCell: 'N22',             // Простой
                absenceCell: 'N23',          // Отсутствия
                percentageCell: 'N25',       // Процент загрузки
                overtimeCell: 'N24',         // Сверхурочные часы
                title: 'Текущий день',       // Резервный заголовок графика
                titleCell: 'A20',            // Заголовок графика из Excel
                subtitleCell: 'A21',         // Подзаголовок (период/даты)
                percentagesRow: 25           // Процент загрузки по управлениям
            },
            {
                id: 'last-week',             // Нарастающий итог по дням за месяц
                loadCell: 'N26',
                totalCell: 'N27',
                idleCell: 'N28',
                absenceCell: 'N29',
                percentageCell: 'N31',
                overtimeCell: 'N30',
                title: 'Нарастающий итог по неделям',
                titleCell: 'A26',
                subtitleCell: 'A27',
                percentagesRow: 31
            },
            {
                id: 'month',                 // Текущий месяц
                loadCell: 'N32',
                totalCell: 'N33',
                idleCell: 'N34',
                absenceCell: 'N35',
                percentageCell: 'N37',
                overtimeCell: 'N36',
                title: 'Текущий месяц',
                titleCell: 'A32',
                subtitleCell: 'A33',
                percentagesRow: 37
            },
            {
                id: 'last-month',            // Прошлый месяц
                loadCell: 'N38',
                totalCell: 'N39',
                idleCell: 'N40',
                absenceCell: 'N41',
                percentageCell: 'N44',
                overtimeCell: 'N42',
                title: 'Прошлый месяц',
                titleCell: 'A38',
                subtitleCell: 'A39',
                percentagesRow: 44
            },
            {
                id: 'next-month',            // Следующий месяц
                loadCell: 'N45',
                totalCell: 'N46',
                idleCell: 'N47',
                absenceCell: 'N48',
                percentageCell: 'N50',
                overtimeCell: 'N49',
                title: 'Следующий месяц',
                titleCell: 'A45',
                subtitleCell: 'A46',
                percentagesRow: 50
            }
        ]
    },

    // Конфигурация для квартальной динамики (тренд по кварталам)
    quarterlyTrend: {
        // Три текущих месяца (прошлый, текущий, следующий)
        currentMonths: [
            {
                labelCell: 'A39',      // Название прошлого месяца
                loadCell: 'N38',       // Часы за прошлый месяц
                totalCell: 'N39',      // Всего часов за прошлый месяц
                percentageCell: 'N44', // Процент загрузки за прошлый месяц
                monthOffset: -1        // Смещение от текущей даты (-1 = прошлый месяц)
            },
            {
                labelCell: 'A33',      // Текущий месяц
                loadCell: 'N32',       // Факт часы
                totalCell: 'N33',      // Всего часов
                percentageCell: 'N37', // Процент загрузки
                monthOffset: 0         // 0 = текущий месяц
            },
            {
                labelCell: 'A46',      // Следующий месяц
                loadCell: 'N45',       // Факт часы
                totalCell: 'N46',      // Всего часов
                percentageCell: 'N50', // Процент загрузки
                monthOffset: 1         // +1 = следующий месяц
            }
        ],
        // Исторические данные (блоки по 5 строк на каждый месяц)
        history: {
            monthLabelColumn: 'A',    // Название месяца
            periodStartColumn: 'A',   // Дата начала периода
            periodEndColumn: 'A',     // Дата окончания периода
            loadColumn: 'N',          // Загрузка в часах
            totalColumn: 'N',         // Всего часов
            percentageColumn: 'N',    // Процент загрузки
            monthNameStartRow: 55,    // Начальная строка для исторических данных
            blockStep: 5,             // Каждый месяц занимает 5 строк (шаг 5)
            loadRowOffset: -1,        // Загрузка на 1 строку выше от строки с названием
            totalRowOffset: 0,        // Всего часов на той же строке
            periodStartRowOffset: 2,  // Дата начала через 2 строки от названия
            periodEndRowOffset: 3,    // Дата окончания через 3 строки от названия
            percentageRowOffset: 3,   // Процент через 3 строки от названия
            maxBlocks: 24             // Максимум 24 блока (месяца) для отображения
        }
    },
    
    // Конфигурация проектов
    projects: {
        // Верхняя таблица со сводкой по статусам и типам проектов
        topTable: {
            titleColumn: 'BL',               // Столбец BL: названия статусов проектов
            titleStartRow: 5,                // Начальная строка статусов (BL5-BL9)
            titleEndRow: 9,                  // Конечная строка статусов
            headersStartCol: 'BM',           // Начальная колонка заголовков типов проектов
            headersEndCol: 'BP',             // Конечная колонка заголовков типов проектов (BM4-BP4)
            headersRow: 4                    // Строка с названиями типов проектов
        },
        // Детальная таблица со списком всех проектов
        detailsTable: {
            startCol: 'AM',                  // Начальная колонка детальной таблицы
            endCol: 'BA',                    // Конечная колонка детальной таблицы
            startRow: 5,                    // Начальная строка данных
            columns: {
                id: 'AN',                    // Колонка AN: ID проекта
                management: 'AO',            // Колонка AO: управление/центр
                projectType: 'AQ',           // Колонка AQ: тип проекта
                status: 'AR',                // Колонка AR: статус проекта
                name: 'AP',                  // Колонка AP: название проекта
                planDate: 'AS',              // Колонка AS: плановая дата
                estimateDate: 'AT',          // Колонка AT: прогнозируемая дата
                dateIndicatorColor: 'AW',    // Колонка AW: цвет индикатора по дате (числовое значение)
                dateIndicator: 'AX',         // Колонка AX: текстовый индикатор по дате
                planHours: 'AU',             // Колонка AU: плановые часы
                actualHours: 'AV',           // Колонка AV: фактические часы
                hoursIndicatorColor: 'AY',   // Колонка AY: цвет индикатора по часам (числовое значение)
                hoursIndicator: 'AZ'         // Колонка AZ: процент остатка по часам
            }
        },
        // Типы проектов и их значения по периодам
        typesRange: {
            titleRow: 4,                     // Строка 4: названия типов проектов
            titleStartCol: 'O',              // Начальная колонка (O4)
            titleEndCol: 'S',                // Конечная колонка (S4) - всего 5 типов проектов
            
            // Периоды и их диапазоны значений (проценты по типам проектов)
            periods: [
                {
                    name: 'Период 1',        // Текущий день
                    subtitleCell: 'A21',     // Ячейка с подзаголовком периода
                    valuesRow: 25,           // Строка 25: значения процентов по типам проектов
                    startCol: 'O',           // Начальная колонка (O24)
                    endCol: 'S'              // Конечная колонка (S24)
                },
                {
                    name: 'Период 2',        // Нарастающий итог за месяц
                    subtitleCell: 'A27',
                    valuesRow: 31,           // Строка 31: значения процентов
                    startCol: 'O',
                    endCol: 'S'
                },
                {
                    name: 'Период 3',        // Текущий месяц
                    subtitleCell: 'A33',
                    valuesRow: 37,           // Строка 37: значения процентов
                    startCol: 'O',
                    endCol: 'S'
                },
                {
                    name: 'Период 4',        // Прошлый месяц
                    subtitleCell: 'A39',
                    valuesRow: 44,           // Строка 44: значения процентов
                    startCol: 'O',
                    endCol: 'S'
                },
                {
                    name: 'Период 5',        // Следующий месяц
                    subtitleCell: 'A46',
                    valuesRow: 50,           // Строка 50: значения процентов
                    startCol: 'O',
                    endCol: 'S'
                }
            ]
        }
    },
    
    // Конфигурация управлений/центров
    managements: {
        nameRow: 4,                          // Строка 4: названия управлений (колонки C-K)
        countRow: 6,                         // Строка 6: количество сотрудников в управлении
        percentageRow: 35,                   // НЕ ИСПОЛЬЗУЕТСЯ (зарезервировано для общих процентов)
        loadHoursRow: 33,                    // НЕ ИСПОЛЬЗУЕТСЯ (зарезервировано для часов загрузки)
        totalHoursRow: 34,                   // НЕ ИСПОЛЬЗУЕТСЯ (зарезервировано для общего количества часов)
        columns: ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']  // Колонки C-K: данные по 9 управлениям/центрам
    }
};

const INDICATOR_COLORS = {
    get GREEN() { return getCssVar('--accent-green', '#84c39a'); },
    get YELLOW() { return getCssVar('--accent-amber-strong', '#d7a44b'); },
    get RED() { return getCssVar('--accent-red-strong', '#d96d61'); },
    get GRAY() { return getCssVar('--accent-gray', '#cbd5e1'); }
};

const RU_MONTH_INDEX = {
    'январь': 0,
    'февраль': 1,
    'март': 2,
    'апрель': 3,
    'май': 4,
    'июнь': 5,
    'июль': 6,
    'август': 7,
    'сентябрь': 8,
    'октябрь': 9,
    'ноябрь': 10,
    'декабрь': 11
};

const RU_MONTH_NAMES = [
    'январь',
    'февраль',
    'март',
    'апрель',
    'май',
    'июнь',
    'июль',
    'август',
    'сентябрь',
    'октябрь',
    'ноябрь',
    'декабрь'
];

const QUARTER_TITLES = ['I квартал', 'II квартал', 'III квартал', 'IV квартал'];

// ============================================================
// УТИЛИТЫ
// ============================================================

function createWorksheetCache(worksheet) {
    const cache = {};
    
    for (const key in worksheet) {
        if (key[0] === '!') continue;
        
        const cell = worksheet[key];
        if (cell && cell.v !== undefined) {
            cache[key] = cell.v;
        }
    }
    
    return cache;
}

function getCachedWorksheetCache(sheetName) {
    if (workbookDataCache.worksheetCellCaches.has(sheetName)) {
        return workbookDataCache.worksheetCellCaches.get(sheetName);
    }

    const worksheet = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
    if (!worksheet) return null;

    const cache = createWorksheetCache(worksheet);
    workbookDataCache.worksheetCellCaches.set(sheetName, cache);
    return cache;
}

function getCssVar(name, fallback = '') {
    const value = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return value || fallback;
}

function getCellValue(cellAddress, cache) {
    if (!cache) return '';
    const value = cache[cellAddress];
    return value !== undefined && value !== null ? value : '';
}

function isValid(value) {
    return value !== undefined && value !== null && value !== '';
}

function formatExcelDate(excelDate) {
    if (!isValid(excelDate)) return '';
    
    try {
        if (typeof excelDate === 'number') {
            if (excelDate <= 0) return '';

            const date = new Date((excelDate - 25569) * 86400 * 1000);
            if (!isNaN(date.getTime())) {
                if (date.getUTCFullYear() <= 1899) return '';
                return date.toLocaleDateString('ru-RU');
            }
        }
        
        if (excelDate instanceof Date) {
            if (excelDate.getFullYear() <= 1899) return '';
            return excelDate.toLocaleDateString('ru-RU');
        }
        
        if (typeof excelDate === 'string') {
            const trimmedDate = excelDate.trim();
            if (!trimmedDate || trimmedDate === '-' || trimmedDate === '—') return '';
            if (/^30[.\/-]12[.\/-]1899$/.test(trimmedDate)) return '';

            const num = parseFloat(trimmedDate);
            if (!isNaN(num) && /^\d+(\.\d+)?$/.test(trimmedDate)) {
                return formatExcelDate(num);
            }
            
            const date = new Date(trimmedDate);
            if (!isNaN(date.getTime())) {
                if (date.getFullYear() <= 1899) return '';
                return date.toLocaleDateString('ru-RU');
            }
            
            return trimmedDate;
        }
        
        return String(excelDate);
    } catch (e) {
        console.warn('Ошибка форматирования даты:', e);
        return '';
    }
}

function formatMonthName(monthIndex) {
    return RU_MONTH_NAMES[monthIndex] || '';
}

function formatQuarterRangeLabel(quarter, year) {
    const quarterIndex = Math.max(1, Math.min(4, quarter)) - 1;
    const startMonth = quarterIndex * 3;
    const endMonth = startMonth + 2;
    return [
        QUARTER_TITLES[quarterIndex],
        `${formatMonthName(startMonth)}-${formatMonthName(endMonth)}`,
        String(year)
    ];
}

function flattenChartLabel(label) {
    return Array.isArray(label) ? label.join(' ') : String(label || '');
}

function excelSerialToDate(value) {
    if (typeof value !== 'number' || !Number.isFinite(value)) return null;
    const utcDays = Math.floor(value - 25569);
    const utcValue = utcDays * 86400;
    const dateInfo = new Date(utcValue * 1000);

    if (isNaN(dateInfo.getTime())) {
        return null;
    }

    return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate());
}

function getReferenceDate() {
    const handbookSheet = workbook && workbook.Sheets ? workbook.Sheets['Справочник'] : null;
    const handbookCell = handbookSheet && handbookSheet.C2 ? handbookSheet.C2.v : null;

    if (typeof handbookCell === 'number') {
        const date = new Date((handbookCell - 25569) * 86400 * 1000);
        if (!isNaN(date.getTime())) {
            return new Date(date.getFullYear(), date.getMonth(), date.getDate());
        }
    }

    if (typeof handbookCell === 'string') {
        const parsed = parseDateValue(handbookCell) || new Date(handbookCell);
        if (!isNaN(parsed?.getTime?.())) {
            return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
        }
    }

    if (fileLastModified instanceof Date && !isNaN(fileLastModified.getTime())) {
        return new Date(fileLastModified.getFullYear(), fileLastModified.getMonth(), fileLastModified.getDate());
    }

    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), now.getDate());
}

function parseRuMonthName(value) {
    if (!isValid(value)) return null;

    if (value instanceof Date && !isNaN(value.getTime())) {
        return value.getMonth();
    }

    const parsedDate = parseDateValue(value);
    if (parsedDate) {
        return parsedDate.getMonth();
    }

    const normalized = String(value)
        .trim()
        .toLowerCase()
        .replace(/ё/g, 'е')
        .split(/\s+/)[0];

    return Object.prototype.hasOwnProperty.call(RU_MONTH_INDEX, normalized)
        ? RU_MONTH_INDEX[normalized]
        : null;
}

function hexToRgba(color, alpha = 1) {
    if (!color) return `rgba(132, 195, 154, ${alpha})`;

    const trimmed = String(color).trim();

    if (trimmed.startsWith('#')) {
        let hex = trimmed.slice(1);
        if (hex.length === 3) {
            hex = hex.split('').map(char => char + char).join('');
        }

        if (hex.length === 6) {
            const r = parseInt(hex.slice(0, 2), 16);
            const g = parseInt(hex.slice(2, 4), 16);
            const b = parseInt(hex.slice(4, 6), 16);
            return `rgba(${r}, ${g}, ${b}, ${alpha})`;
        }
    }

    const rgbMatch = trimmed.match(/rgba?\(([^)]+)\)/i);
    if (rgbMatch) {
        const parts = rgbMatch[1].split(',').map(part => part.trim());
        const [r, g, b] = parts;
        return `rgba(${r}, ${g}, ${b}, ${alpha})`;
    }

    return trimmed;
}

function createRelativeMonthRecord(config, referenceDate) {
    const label = getCellValue(config.labelCell, worksheetCache);
    const rawLoad = getCellValue(config.loadCell, worksheetCache);
    const rawTotal = getCellValue(config.totalCell, worksheetCache);
    const rawPercentage = getCellValue(config.percentageCell, worksheetCache);

    if (!isValid(label) && !isValid(rawLoad) && !isValid(rawTotal) && !isValid(rawPercentage)) {
        return null;
    }

    const date = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + config.monthOffset, 1);
    const loadHours = parseFloat(rawLoad);
    const totalHours = parseFloat(rawTotal);
    const percentage = parseFloat(rawPercentage);

    return {
        year: date.getFullYear(),
        monthIndex: date.getMonth(),
        label: String(label || '').trim(),
        loadHours: Number.isFinite(loadHours) ? loadHours : 0,
        totalHours: Number.isFinite(totalHours) ? totalHours : 0,
        percentage: Number.isFinite(percentage)
            ? percentage
            : ((Number.isFinite(loadHours) && Number.isFinite(totalHours) && totalHours > 0)
                ? (loadHours / totalHours) * 100
                : null),
        sourcePriority: 2
    };
}

function createHistoricalMonthRecord(monthNameRow, config) {
    const monthLabel = getCellValue(`${config.monthLabelColumn}${monthNameRow}`, worksheetCache);
    const startDateRaw = getCellValue(`${config.periodStartColumn}${monthNameRow + config.periodStartRowOffset}`, worksheetCache);
    const endDateRaw = getCellValue(`${config.periodEndColumn}${monthNameRow + config.periodEndRowOffset}`, worksheetCache);
    const rawLoad = getCellValue(`${config.loadColumn}${monthNameRow + config.loadRowOffset}`, worksheetCache);
    const rawTotal = getCellValue(`${config.totalColumn}${monthNameRow + config.totalRowOffset}`, worksheetCache);
    const rawPercentage = getCellValue(`${config.percentageColumn}${monthNameRow + config.percentageRowOffset}`, worksheetCache);

    if (!isValid(monthLabel) && !isValid(startDateRaw) && !isValid(endDateRaw) && !isValid(rawLoad) && !isValid(rawTotal) && !isValid(rawPercentage)) {
        return null;
    }

    const explicitDate = parseDateValue(endDateRaw) || parseDateValue(monthLabel);
    let year = null;
    let monthIndex = null;

    if (explicitDate) {
        year = explicitDate.getFullYear();
        monthIndex = explicitDate.getMonth();
    } else {
        monthIndex = parseRuMonthName(monthLabel);
    }

    if (monthIndex === null) {
        return null;
    }

    if (year === null) {
        year = getReferenceDate().getFullYear();
    }

    const loadHours = parseFloat(rawLoad);
    const totalHours = parseFloat(rawTotal);
    const percentage = parseFloat(rawPercentage);

    return {
        year,
        monthIndex,
        label: formatMonthName(monthIndex),
        periodStart: isValid(startDateRaw) ? String(startDateRaw).trim() : '',
        periodEnd: isValid(endDateRaw) ? String(endDateRaw).trim() : '',
        loadHours: Number.isFinite(loadHours) ? loadHours : 0,
        totalHours: Number.isFinite(totalHours) ? totalHours : 0,
        percentage: Number.isFinite(percentage)
            ? percentage
            : ((Number.isFinite(loadHours) && Number.isFinite(totalHours) && totalHours > 0)
                ? (loadHours / totalHours) * 100
                : null),
        sourcePriority: 1
    };
}

function getQuarterlyTrendMonths() {
    const referenceDate = getReferenceDate();
    const records = EXCEL_CONFIG.quarterlyTrend.currentMonths.map(config =>
        createRelativeMonthRecord(config, referenceDate)
    );
    const historyConfig = EXCEL_CONFIG.quarterlyTrend.history;

    for (let index = 0; index < historyConfig.maxBlocks; index++) {
        const monthNameRow = historyConfig.monthNameStartRow + (index * historyConfig.blockStep);
        records.push(createHistoricalMonthRecord(monthNameRow, historyConfig));
    }

    const uniqueRecords = new Map();
    records.filter(Boolean).forEach(record => {
        const key = `${record.year}-${record.monthIndex}`;
        const existing = uniqueRecords.get(key);
        if (!existing || record.sourcePriority > existing.sourcePriority) {
            uniqueRecords.set(key, record);
        }
    });

    return Array.from(uniqueRecords.values()).sort((a, b) => {
        const yearDiff = a.year - b.year;
        return yearDiff !== 0 ? yearDiff : a.monthIndex - b.monthIndex;
    });
}

function getQuarterlyTrendData() {
    return buildQuarterlyTrendData(getQuarterlyTrendMonths());
}

function buildQuarterlyTrendData(months) {
    const quarterMap = new Map();

    months.forEach(record => {
        const quarter = Math.floor(record.monthIndex / 3) + 1;
        const key = `${record.year}-Q${quarter}`;

        if (!quarterMap.has(key)) {
            quarterMap.set(key, {
                key,
                year: record.year,
                quarter,
                label: formatQuarterRangeLabel(quarter, record.year),
                loadHours: 0,
                totalHours: 0,
                months: [],
                percentages: [],
                monthIndexes: []
            });
        }

        const current = quarterMap.get(key);
        current.loadHours += record.loadHours;
        current.totalHours += record.totalHours;
        current.months.push(record.label);
        current.monthIndexes.push(record.monthIndex);
        if (Number.isFinite(record.percentage)) {
            current.percentages.push(record.percentage);
        }
    });

    return Array.from(quarterMap.values()).sort((a, b) => {
        const yearDiff = a.year - b.year;
        return yearDiff !== 0 ? yearDiff : a.quarter - b.quarter;
    }).map(item => {
        const percentage = item.totalHours > 0
            ? (item.loadHours / item.totalHours) * 100
            : (item.percentages.length > 0
                ? item.percentages.reduce((sum, value) => sum + value, 0) / item.percentages.length
                : 0);

        return {
            ...item,
            label: formatQuarterRangeLabel(item.quarter, item.year),
            percentage: clampQuarterlyTrendPercentage(percentage)
        };
    });
}

function clampQuarterlyTrendPercentage(value) {
    if (!Number.isFinite(value)) return 0;
    return Math.min(Math.max(value, 0), 100);
}

function getTechBlockQuarterlyTrendDataFromTables() {
    const actualHoursByMonth = getTechBlockActualHoursByMonthFromLoadTable();
    const monthIntervals = getTechBlockQuarterlyTrendMonthIntervals(actualHoursByMonth);

    if (monthIntervals.length === 0) return [];

    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);
    const capacityByMonth = new Map();

    capacityByCenterMonth.forEach(monthMap => {
        monthMap.forEach((value, monthKey) => {
            addMonthMapValue(capacityByMonth, monthKey, value.availableHours || 0);
        });
    });

    const months = monthIntervals
        .map(interval => {
            const monthKey = getMonthKey(interval.start);
            const loadHours = actualHoursByMonth.get(monthKey) || 0;
            const totalHours = capacityByMonth.get(monthKey) || 0;

            return {
                year: interval.start.getFullYear(),
                monthIndex: interval.start.getMonth(),
                label: formatMonthName(interval.start.getMonth()),
                loadHours,
                totalHours,
                percentage: totalHours > 0 ? (loadHours / totalHours) * 100 : null,
                sourcePriority: 2
            };
        })
        .filter(month => month.loadHours > 0 || month.totalHours > 0);

    return buildQuarterlyTrendData(months);
}

function getCenterQuarterlyTrendDataFromTables(centerKey) {
    if (!centerKey) return [];

    const monthIntervals = getQuarterlyTrendDisplayMonthIntervals();
    if (monthIntervals.length === 0) return [];

    const actualHoursByMonth = getSpentProjectHoursByMonthForCenter(centerKey);
    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);
    const centerCapacityByMonth = capacityByCenterMonth.get(centerKey) || new Map();

    const months = monthIntervals.map(interval => {
        const monthKey = getMonthKey(interval.start);
        const loadHours = actualHoursByMonth.get(monthKey) || 0;
        const totalHours = centerCapacityByMonth.get(monthKey)?.availableHours || 0;

        return {
            year: interval.start.getFullYear(),
            monthIndex: interval.start.getMonth(),
            label: formatMonthName(interval.start.getMonth()),
            loadHours,
            totalHours,
            percentage: totalHours > 0 ? (loadHours / totalHours) * 100 : null,
            sourcePriority: 2
        };
    });

    return buildQuarterlyTrendData(months).filter(quarter => quarter.loadHours > 0 || quarter.totalHours > 0);
}

function getQuarterlyTrendDisplayMonthIntervals(referenceDate = getReferenceDate()) {
    const quarterStartMonth = Math.floor(referenceDate.getMonth() / 3) * 3;
    const endDate = new Date(referenceDate.getFullYear(), quarterStartMonth + 3, 1);
    const startDate = new Date(referenceDate.getFullYear(), quarterStartMonth - ((QUARTERLY_TREND_DISPLAY_LIMIT - 1) * 3), 1);

    return getPlanFactMonthIntervals(startDate, endDate);
}

function getQuarterlyTrendDisplayQuarters(quarters, options = {}) {
    const displayQuarters = (Array.isArray(quarters) ? quarters : [])
        .slice(-QUARTERLY_TREND_DISPLAY_LIMIT)
        .map(quarter => ({
            ...quarter,
            percentage: clampQuarterlyTrendPercentage(quarter.percentage)
        }));

    if (options.holdLastFutureDrop === true && displayQuarters.length >= 2) {
        const lastIndex = displayQuarters.length - 1;
        const previous = displayQuarters[lastIndex - 1];
        const current = displayQuarters[lastIndex];

        if (
            isQuarterlyTrendQuarterFutureAffected(current)
            && Number.isFinite(previous.percentage)
            && Number.isFinite(current.percentage)
            && current.percentage < previous.percentage
        ) {
            displayQuarters[lastIndex] = {
                ...current,
                originalPercentage: current.percentage,
                percentage: previous.percentage,
                isHeldFromFutureDrop: true
            };
        }
    }

    return displayQuarters;
}

function isQuarterlyTrendQuarterFutureAffected(quarter) {
    if (!quarter || !Number.isFinite(quarter.year) || !Number.isFinite(quarter.quarter)) return false;

    const nextMonthStart = getNextMonthStartDate();
    const quarterEnd = new Date(quarter.year, quarter.quarter * 3, 1);

    return quarterEnd > nextMonthStart;
}

function getTechBlockQuarterlyTrendMonthIntervals(actualHoursByMonth) {
    const referenceDate = getReferenceDate();
    const nextMonthStart = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 1);
    const actualMonthStarts = Array.from(actualHoursByMonth.keys())
        .map(parseMonthKeyToDate)
        .filter(Boolean)
        .sort((a, b) => a - b);
    const fallbackRange = getPlanFactDisplayRange();
    const firstActualMonth = actualMonthStarts[0] || fallbackRange.startDate;
    const startDate = new Date(firstActualMonth.getFullYear(), firstActualMonth.getMonth(), 1);
    const endDate = nextMonthStart < fallbackRange.endDate ? nextMonthStart : fallbackRange.endDate;

    if (!(startDate instanceof Date) || !(endDate instanceof Date) || startDate >= endDate) {
        return [];
    }

    return getPlanFactMonthIntervals(startDate, endDate);
}

function getTechBlockActualHoursByMonthFromLoadTable() {
    const result = new Map();
    const context = getWorksheetHeaderContext('Загрузка_ТехБлок');
    if (!context) return result;

    const employeeLookup = getTechBlockEmployeeLookup();
    const nextMonthStart = getNextMonthStartDate();

    for (let rowIndex = context.range.s.r + 1; rowIndex <= context.range.e.r; rowIndex++) {
        const id = normalizeProjectId(getWorksheetContextValue(context, rowIndex, ['ID']));
        if (!id) continue;

        const hours = parseNumericValue(getWorksheetContextValue(context, rowIndex, ['Часы'])) || 0;
        if (hours <= 0) continue;

        const date = parseDateValue(getWorksheetContextValue(context, rowIndex, ['Дата']));
        if (!date || date >= nextMonthStart) continue;

        const fio = normalizePersonName(getWorksheetContextValue(context, rowIndex, ['ФИО']));
        const employee = employeeLookup.get(fio);
        const rowTerminationDate = parseDateValue(getWorksheetContextValue(context, rowIndex, ['Дата увольнения']));
        const terminationDate = rowTerminationDate || employee?.terminationDate || null;
        const isAdmin = isAdminPersonnelValue(getWorksheetContextValue(context, rowIndex, ['Административный персонал'])) || Boolean(employee?.isAdmin);
        const absence = getWorksheetContextValue(context, rowIndex, ['Отсутствия']);
        const isAbsence = isProjectAbsenceValue(absence);
        const isForecast = isForecastLoadValue(getWorksheetContextValue(context, rowIndex, ['Прогноз']));

        if (isAdmin || isAbsence || isForecast || !isEmployeeActiveOnDate(date, terminationDate)) continue;

        addMonthMapValue(result, getMonthKey(date), hours);
    }

    return result;
}

function getEmployeesList(column, startRow, cache, maxRows = 50) {
    const employees = [];
    
    for (let row = startRow; row < startRow + maxRows; row++) {
        const cellAddress = column + row;
        const name = getCellValue(cellAddress, cache);
        
        if (!isValid(name) || String(name).trim() === '') break;
        
        employees.push(String(name).trim());
    }
    
    return employees;
}

function getColumnRange(startCol, endCol) {
    const columns = [];

    function nextColumn(col) {
        let result = '';
        let carry = 1;

        for (let i = col.length - 1; i >= 0; i--) {
            let charCode = col.charCodeAt(i) - 65 + carry;

            if (charCode === 26) {
                charCode = 0;
                carry = 1;
            } else {
                carry = 0;
            }

            result = String.fromCharCode(65 + charCode) + result;
        }

        if (carry === 1) {
            result = 'A' + result;
        }

        return result;
    }

    let current = startCol;

    while (true) {
        columns.push(current);
        if (current === endCol) break;
        current = nextColumn(current);
    }

    return columns;
}

function getColumnIndexFromLetter(columnLetter) {
    const letters = String(columnLetter || '').trim().toUpperCase();
    if (!/^[A-Z]+$/.test(letters)) return -1;

    let index = 0;
    for (const letter of letters) {
        index = index * 26 + (letter.charCodeAt(0) - 64);
    }

    return index - 1;
}

function getWorksheetCellByColumnLetter(context, rowIndex, columnLetter) {
    if (!context || typeof XLSX === 'undefined') return '';

    const columnIndex = getColumnIndexFromLetter(columnLetter);
    if (columnIndex < 0) return '';

    const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
    return getCellValue(address, context.cache);
}

function escapeHtml(text) {
    if (text === undefined || text === null) return '';
    const div = document.createElement('div');
    div.textContent = String(text);
    return div.innerHTML;
}

function parseNumericValue(value) {
    if (!isValid(value)) return null;
    if (typeof value === 'number') return Number.isFinite(value) ? value : null;

    const normalized = String(value)
        .trim()
        .replace(/\s+/g, '')
        .replace(',', '.');
    const parsed = parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : null;
}

function formatProjectWord(count) {
    const absCount = Math.abs(count) % 100;
    const lastDigit = absCount % 10;

    if (absCount > 10 && absCount < 20) return 'проектов';
    if (lastDigit === 1) return 'проект';
    if (lastDigit >= 2 && lastDigit <= 4) return 'проекта';
    return 'проектов';
}

function formatEmployeeWord(count) {
    const absCount = Math.abs(count) % 100;
    const lastDigit = absCount % 10;

    if (absCount > 10 && absCount < 20) return 'сотрудников';
    if (lastDigit === 1) return 'сотрудник';
    if (lastDigit >= 2 && lastDigit <= 4) return 'сотрудника';
    return 'сотрудников';
}

function formatCenterWord(count) {
    const absCount = Math.abs(count) % 100;
    const lastDigit = absCount % 10;

    if (absCount > 10 && absCount < 20) return 'центров';
    if (lastDigit === 1) return 'центр';
    if (lastDigit >= 2 && lastDigit <= 4) return 'центра';
    return 'центров';
}

function formatNotificationMetric(value, valueType) {
    const numericValue = parseNumericValue(value);
    if (numericValue === null) return '';

    const formatted = Math.round(numericValue).toLocaleString('ru-RU');
    if (valueType === 'hours') {
        return `${formatted} ч.`;
    }

    return `${formatted} ${formatProjectWord(numericValue)}`;
}

function hasNotificationMetricValue(value) {
    const numericValue = parseNumericValue(value);
    if (numericValue !== null) {
        return numericValue > 0;
    }

    return isValid(value) && String(value).trim() !== '';
}

function getNotificationData() {
    const rows = EXCEL_CONFIG.notification.rows.map(config => ({
        key: config.key,
        label: String(getCellValue(config.labelCell, worksheetCache) || '').trim(),
        value: getCellValue(config.valueCell, worksheetCache),
        valueType: config.valueType
    }));

    const fallbackText = String(getCellValue(EXCEL_CONFIG.notification.fallbackCell, worksheetCache) || '').trim();
    const entries = [];

    const deadline = rows.find(row => row.key === 'deadline');
    if (deadline && deadline.label && hasNotificationMetricValue(deadline.value)) {
        entries.push({
            label: deadline.label,
            valueText: formatNotificationMetric(deadline.value, deadline.valueType)
        });
    }

    const hoursProjects = rows.find(row => row.key === 'hoursProjects');
    const hoursTotal = rows.find(row => row.key === 'hoursTotal');
    const hasHoursProjects = hoursProjects && hasNotificationMetricValue(hoursProjects.value);
    const hasHoursTotal = hoursTotal && hasNotificationMetricValue(hoursTotal.value);

    if ((hoursProjects && hoursProjects.label) || (hoursTotal && hoursTotal.label)) {
        const label = (hoursProjects && hoursProjects.label) || (hoursTotal && hoursTotal.label) || 'Вышли по часам';
        const valueParts = [];

        if (hasHoursProjects) {
            valueParts.push(formatNotificationMetric(hoursProjects.value, 'projects'));
        }

        if (hasHoursTotal) {
            valueParts.push(formatNotificationMetric(hoursTotal.value, 'hours'));
        }

        if (valueParts.length > 0) {
            entries.push({
                label,
                valueText: valueParts.join(' / ')
            });
        }
    }

    const missingPlan = rows.find(row => row.key === 'missingPlan');
    if (missingPlan && missingPlan.label && hasNotificationMetricValue(missingPlan.value)) {
        entries.push({
            label: missingPlan.label,
            valueText: formatNotificationMetric(missingPlan.value, missingPlan.valueType)
        });
    }

    return {
        entries,
        hasNotification: entries.length > 0 || fallbackText !== '',
        attentionCount: entries.length > 0 ? entries.length : (fallbackText ? 1 : 0),
        fallbackText
    };
}

// Подсветка найденного текста. Возвращает экранированный HTML с <mark> вокруг совпадений.
// Пустой needle — просто escapeHtml. Регистронезависимо.
function highlightMatch(text, needle) {
    if (!text) return '';
    const safe = escapeHtml(text);
    if (!needle) return safe;
    const needleSafe = escapeHtml(needle);
    if (!needleSafe) return safe;

    // Разбиваем по совпадению (case-insensitive) и собираем обратно с <mark>
    const pattern = new RegExp(needleSafe.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
    return safe.replace(pattern, m => `<mark class="nkitk-mark">${m}</mark>`);
}

// ============================================================
// ФУНКЦИИ ДЛЯ РАБОТЫ С ДАТАМИ (УЛУЧШЕННЫЕ)
// ============================================================

function isValidDate(value) {
    if (value === null || value === undefined || value === '') return false;

    // Работаем только со строками — числа/даты пусть отдельно обрабатываются
    if (typeof value !== 'string') return false;

    const datePattern = /^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/;
    const match = value.match(datePattern);
    if (match) {
        const day = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        const year = parseInt(match[3], 10);

        if (day >= 1 && day <= 31 && month >= 1 && month <= 12 && year >= 1900 && year <= 2100) {
            return true;
        }
    }

    if (value.length <= 5) return false;
    if (/^\d+([.,]\d+)?$/.test(value)) return false;

    const date = new Date(value);
    return !isNaN(date.getTime());
}

function parseDateValue(value) {
    if (value === null || value === undefined || value === '') return null;

    if (typeof value === 'number') {
        if (value <= 0) return null;
        return excelSerialToDate(value);
    }

    if (value instanceof Date) {
        if (isNaN(value.getTime()) || value.getFullYear() <= 1899) return null;
        return value;
    }

    const str = String(value);

    if (/^\d+([.,]\d+)?$/.test(str)) {
        const serial = parseFloat(str.replace(',', '.'));
        if (serial <= 0) return null;
        return excelSerialToDate(serial);
    }

    const datePattern = /^(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})$/;
    const match = str.match(datePattern);
    if (match) {
        const day = parseInt(match[1], 10);
        const month = parseInt(match[2], 10) - 1;
        const year = parseInt(match[3], 10);
        const date = new Date(year, month, day);
        if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
            return date;
        }
        return null;
    }

    const isoLocalPattern = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T]\d{1,2}:\d{2}(?::\d{2})?)?$/;
    const isoLocalMatch = str.match(isoLocalPattern);
    if (isoLocalMatch) {
        const year = parseInt(isoLocalMatch[1], 10);
        const month = parseInt(isoLocalMatch[2], 10) - 1;
        const day = parseInt(isoLocalMatch[3], 10);
        const date = new Date(year, month, day);
        if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
            return date;
        }
        return null;
    }

    const date = new Date(str);
    if (isNaN(date.getTime()) || date.getFullYear() <= 1899) return null;
    return date;
}

function getTodayCalendarDate() {
    const today = new Date();
    return new Date(today.getFullYear(), today.getMonth(), today.getDate());
}

function isSameCalendarDate(dateA, dateB) {
    if (!(dateA instanceof Date) || !(dateB instanceof Date)) return false;
    if (isNaN(dateA.getTime()) || isNaN(dateB.getTime())) return false;

    return dateA.getFullYear() === dateB.getFullYear()
        && dateA.getMonth() === dateB.getMonth()
        && dateA.getDate() === dateB.getDate();
}

function normalizeAbsenceCode(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
}

function isDateColumn(colIndex, data) {
    if (data.length < 2) return false;

    let dateCount = 0;
    let numberCount = 0;
    const checkRows = Math.min(20, data.length - 1);

    for (let i = 1; i <= checkRows; i++) {
        const raw = data[i][colIndex];
        if (raw === null || raw === undefined || raw === '') continue;

        const str = String(raw);

        if (isValidDate(str)) {
            dateCount++;
        }

        if (/^\d+([.,]\d+)?$/.test(str)) {
            numberCount++;
        }
    }

    const totalChecked = dateCount + numberCount;
    if (totalChecked === 0) return false;

    return dateCount > numberCount &&
           dateCount / totalChecked > 0.4 &&
           numberCount / totalChecked < 0.3;
}

function sortByDate(data, colIndex, order = 'asc') {
    const headers = data[0];
    const dataRows = data.slice(1);

    dataRows.sort((a, b) => {
        const dateA = parseDateValue(a[colIndex]);
        const dateB = parseDateValue(b[colIndex]);

        if (!dateA && !dateB) return 0;
        if (!dateA) return 1;
        if (!dateB) return -1;

        return order === 'asc' ? dateA - dateB : dateB - dateA;
    });

    return [headers, ...dataRows];
}

function sortByText(data, colIndex, order = 'asc') {
    const headers = data[0];
    const dataRows = data.slice(1);
    const collator = new Intl.Collator('ru', { numeric: true, sensitivity: 'base' });

    dataRows.sort((a, b) => {
        const aVal = a[colIndex];
        const bVal = b[colIndex];
        const aEmpty = aVal === null || aVal === undefined || aVal === '';
        const bEmpty = bVal === null || bVal === undefined || bVal === '';

        if (aEmpty && bEmpty) return 0;
        if (aEmpty) return 1;
        if (bEmpty) return -1;

        const cmp = collator.compare(String(aVal), String(bVal));
        return order === 'asc' ? cmp : -cmp;
    });

    return [headers, ...dataRows];
}

// ============================================================
// ФУНКЦИИ ДЛЯ НКИТК
// ============================================================

function getUniqueValuesFromColumn(data, colIndex) {
    const uniqueValues = new Set();

    for (let i = 1; i < data.length; i++) {
        const raw = data[i][colIndex];
        if (raw === null || raw === undefined || raw === '') continue;
        const str = String(raw).trim();
        if (str !== '') uniqueValues.add(str);
    }

    return Array.from(uniqueValues).sort((a, b) => a.localeCompare(b, 'ru'));
}

function renderNKiTKTable() {
    if (!worksheetCache) {
        console.warn('Нет данных для отображения таблицы НКиТК');
        const container = document.getElementById('nkitk-table-container');
        if (container) {
            container.innerHTML = '<p>Нет данных для отображения</p>';
        }
        return;
    }
    
    log('Рендерим таблицу НКиТК (столбцы A:G)...');
    
    const COLUMN_COUNT = 7;
    const MAX_ROWS_SCAN = 1000;
    const rawData = [];
    let maxRow = 0;

    // Одним проходом по всем колонкам находим последнюю строку с данными.
    // Сканируем снизу вверх — это быстрее, чем полный перебор.
    outer: for (let row = MAX_ROWS_SCAN; row >= 1; row--) {
        for (let col = 0; col < COLUMN_COUNT; col++) {
            const columnLetter = String.fromCharCode(65 + col);
            const value = getCellValue(columnLetter + row, worksheetCache);

            if (isValid(value) &&
                String(value).trim() !== '' &&
                !String(value).startsWith('#') &&
                !(typeof value === 'string' && value.includes('#'))) {
                maxRow = row;
                break outer;
            }
        }
    }
    
    log(`Последняя строка с данными: ${maxRow}`);
    
    if (maxRow === 0) {
        const container = document.getElementById('nkitk-table-container');
        if (container) {
            container.innerHTML = '<p>Нет данных для отображения</p>';
        }
        return;
    }
    
    for (let row = 1; row <= maxRow; row++) {
    const rowData = [];
    let hasRealData = false;
    
    for (let col = 0; col < COLUMN_COUNT; col++) {
        const columnLetter = String.fromCharCode(65 + col);
        const cellAddress = columnLetter + row;
        let value = getCellValue(cellAddress, worksheetCache);
        
        const isValidValue = isValid(value) && 
                            String(value).trim() !== '' && 
                            !String(value).startsWith('#') &&
                            !(typeof value === 'string' && value.includes('#'));
        
        if (isValidValue) {
            hasRealData = true;
            // Пропускаем преобразование дат для группового столбца (col === 2, это столбец C)
            // Групповой столбец содержит названия (НК, ТК и т.д.), их не нужно преобразовывать в даты
            const isGroupColumn = (col === 2);
            if (!isGroupColumn && typeof value === 'number' && value > 25569 && value < 73050) {
                value = formatExcelDate(value);
            }
            rowData.push(String(value).trim());
        } else {
            rowData.push('');
        }
    }
    
    if (hasRealData) {
        rawData.push(rowData);
    } else {
            let hasDataAfter = false;
            for (let nextRow = row + 1; nextRow <= Math.min(row + 10, maxRow); nextRow++) {
                for (let col = 0; col < COLUMN_COUNT; col++) {
                    const columnLetter = String.fromCharCode(65 + col);
                    const nextValue = getCellValue(columnLetter + nextRow, worksheetCache);
                    if (isValid(nextValue) && String(nextValue).trim() !== '' && !String(nextValue).startsWith('#')) {
                        hasDataAfter = true;
                        break;
                    }
                }
                if (hasDataAfter) break;
            }
            
            if (!hasDataAfter && rawData.length > 0) {
                log(`Остановка на пустой строке ${row}`);
                break;
            }
        }
    }
    
    log(`Собрано строк данных: ${rawData.length}`);
    
    if (rawData.length === 0) {
        const container = document.getElementById('nkitk-table-container');
        if (container) {
            container.innerHTML = '<p>Нет данных для отображения</p>';
        }
        return;
    }
    
    nkitkTableData = rawData;
    const filteredData = applyFilters(nkitkTableData);
    
    const totalRows = filteredData.length - 1;
    const originalRows = nkitkTableData.length - 1;
    
    function getGroupColorsByColumn(data, groupColumnIndex) {
        const groupColors = new Map();
        let currentColor = 0;
        let lastValue = null;
        
        for (let i = 1; i < data.length; i++) {
            const currentValue = data[i][groupColumnIndex];
            if (currentValue !== lastValue) {
                currentColor = currentColor === 0 ? 1 : 0;
                lastValue = currentValue;
            }
            groupColors.set(i, currentColor);
        }
        
        return groupColors;
    }
    
    const rowColors = getGroupColorsByColumn(filteredData, 2);

    const hasActiveFilters = Object.keys(nkitkFilters).length > 0
        || (nkitkGlobalSearch && nkitkGlobalSearch.trim() !== '')
        || !!nkitkSort;

    const globalSearchValue = escapeHtml(nkitkGlobalSearch || '');

    // Верхняя панель: глобальный поиск + одна строка со статистикой и кнопкой сброса
    let html = `
        <div class="nkitk-toolbar">
            <div class="nkitk-search">
                <span class="nkitk-search__icon">🔍</span>
                <input type="text"
                       id="nkitk-global-search"
                       class="nkitk-search__input"
                       placeholder="Поиск по всем столбцам..."
                       value="${globalSearchValue}">
                ${nkitkGlobalSearch ? '<button type="button" class="nkitk-search__clear" id="nkitk-global-search-clear" title="Очистить поиск">✖</button>' : ''}
            </div>
            <div class="nkitk-toolbar__stats">
                <span class="nkitk-toolbar__count">
                    Показано <strong>${totalRows}</strong> из <strong>${originalRows}</strong>
                    ${totalRows !== originalRows ? `<span class="nkitk-toolbar__hidden">(скрыто ${originalRows - totalRows})</span>` : ''}
                </span>
                ${hasActiveFilters ? `<button type="button" class="nkitk-btn nkitk-btn--danger nkitk-btn--small" id="nkitk-clear-all">Сбросить всё</button>` : ''}
            </div>
        </div>
        <div id="active-filters-display" class="nkitk-filters-bar__active"></div>
    `;

    html += `
        <div class="nkitk-table-wrapper">
            <table class="nkitk-data-table nkitk-data-table--sticky">
                <thead>
    `;

    if (filteredData.length > 0) {
        const headers = filteredData[0];
        html += '<tr>';
        for (let col = 0; col < headers.length; col++) {
            const headerValue = headers[col] || String.fromCharCode(65 + col);
            const colFilter = nkitkFilters[col];
            const hasColFilter = !!(colFilter && (
                (Array.isArray(colFilter.checkedValues) && colFilter.checkedValues.length > 0) ||
                (colFilter.dateFilter && colFilter.dateFilter !== 'all') ||
                colFilter.customRange
            ));
            const isGroupColumn = (col === 2);
            const isSorted = nkitkSort && nkitkSort.col === col;
            const sortIndicator = isSorted
                ? (nkitkSort.order === 'asc' ? '↑' : '↓')
                : '↕';

            const thClasses = ['nkitk-th'];
            if (isGroupColumn) thClasses.push('nkitk-data-table__cell--group');
            if (isSorted) thClasses.push('nkitk-th--sorted');

            const filterBtnClass = hasColFilter ? 'nkitk-filter-btn nkitk-filter-btn--active' : 'nkitk-filter-btn';

            html += `
                <th class="${thClasses.join(' ')}">
                    <div class="nkitk-data-table__header-inner">
                        <button type="button" class="nkitk-th__sort" data-col="${col}" title="Сортировать">
                            <span class="nkitk-th__label">${escapeHtml(headerValue)}</span>
                            <span class="nkitk-th__arrow">${sortIndicator}</span>
                        </button>
                        <button type="button" class="${filterBtnClass}" data-col="${col}" data-header="${escapeHtml(headerValue)}" title="Фильтр">
                            ▼
                        </button>
                    </div>
                </th>
            `;
        }
        html += '</tr></thead><tbody>';

        const globalNeedle = (nkitkGlobalSearch || '').trim();

        for (let row = 1; row < filteredData.length; row++) {
            const rowClass = rowColors.get(row) === 0
                ? 'nkitk-data-table__row--default'
                : 'nkitk-data-table__row--alt';

            html += `<tr class="${rowClass}">`;
            for (let col = 0; col < filteredData[row].length; col++) {
                const cellValue = filteredData[row][col] || '';
                const colFilter = nkitkFilters[col];
                const isFilteredCell = !!(colFilter && Array.isArray(colFilter.checkedValues)
                    && colFilter.checkedValues.length > 0
                    && colFilter.checkedValues.some(v => String(v).toLowerCase() === String(cellValue).toLowerCase()));
                const isGroupColumn = (col === 2);

                const classes = [];
                if (isFilteredCell) classes.push('nkitk-data-table__cell--filtered');
                if (isGroupColumn) classes.push('nkitk-data-table__cell--group');
                const classAttr = classes.length ? ` class="${classes.join(' ')}"` : '';

                html += `<td${classAttr}>${highlightMatch(String(cellValue), globalNeedle)}</td>`;
            }
            html += '</tr>';
        }
        html += '</tbody>';
    } else {
        html += `<tr><td colspan="${COLUMN_COUNT}" class="nkitk-data-table__empty">Ничего не найдено по заданным фильтрам</td></tr></tbody>`;
    }

    html += '</table></div>';

    // Контейнер для popover'а фильтра (создаём один раз)
    if (!document.getElementById('nkitk-filter-popover')) {
        const popoverHtml = `<div id="nkitk-filter-popover" class="nkitk-popover" hidden></div>`;
        document.body.insertAdjacentHTML('beforeend', popoverHtml);
    }
    
    const container = document.getElementById('nkitk-table-container');
    if (container) {
        container.innerHTML = html;
        updateActiveFiltersDisplay();
        attachNKiTKHandlers();
        log(`Таблица НКиТК отрисована: ${filteredData.length} строк`);
    }
}

// Обработчики: поиск, сортировка по клику, кнопки фильтра, сброс
function attachNKiTKHandlers() {
    // Глобальный поиск с debounce 200 мс
    const searchInput = document.getElementById('nkitk-global-search');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            const value = e.target.value;
            if (globalSearchDebounceId) clearTimeout(globalSearchDebounceId);
            globalSearchDebounceId = setTimeout(() => {
                nkitkGlobalSearch = value;
                const caretPos = searchInput.selectionStart;
                renderNKiTKTable();
                // Возвращаем фокус и позицию курсора
                const newInput = document.getElementById('nkitk-global-search');
                if (newInput) {
                    newInput.focus();
                    try { newInput.setSelectionRange(caretPos, caretPos); } catch (e) {}
                }
            }, 200);
        });
    }

    const searchClear = document.getElementById('nkitk-global-search-clear');
    if (searchClear) {
        searchClear.addEventListener('click', () => {
            nkitkGlobalSearch = '';
            renderNKiTKTable();
        });
    }

    // Кнопка "Сбросить всё"
    const clearAll = document.getElementById('nkitk-clear-all');
    if (clearAll) {
        clearAll.addEventListener('click', () => window.clearAllFilters());
    }

    // Сортировка кликом по заголовку (кнопка-обёртка внутри th)
    document.querySelectorAll('.nkitk-th__sort').forEach(btn => {
        btn.addEventListener('click', () => {
            const col = parseInt(btn.dataset.col, 10);
            if (!nkitkSort || nkitkSort.col !== col) {
                nkitkSort = { col, order: 'asc' };
            } else if (nkitkSort.order === 'asc') {
                nkitkSort = { col, order: 'desc' };
            } else {
                nkitkSort = null;
            }
            renderNKiTKTable();
        });
    });

    // Кнопка фильтра — открываем popover
    document.querySelectorAll('.nkitk-filter-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const col = parseInt(btn.dataset.col, 10);
            const header = btn.dataset.header || '';
            openFilterPopover(btn, col, header);
        });
    });
}

function applyFilters(data) {
    const hasColumnFilters = Object.keys(nkitkFilters).length > 0;
    const hasGlobal = nkitkGlobalSearch && nkitkGlobalSearch.trim() !== '';

    if (!hasColumnFilters && !hasGlobal && !nkitkSort) {
        return data;
    }

    const globalNeedle = hasGlobal ? nkitkGlobalSearch.trim().toLowerCase() : '';

    let filteredData = [data[0]];
    
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        let includeRow = true;
        
        for (const [colIndex, filter] of Object.entries(nkitkFilters)) {
            const col = parseInt(colIndex);
            
            if (filter.type === 'date') {
                const dateValue = parseDateValue(row[col]);
                if (!dateValue) {
                    includeRow = false;
                    break;
                }
                
                if (filter.dateFilter && filter.dateFilter !== 'all') {
                    const today = new Date();
                    today.setHours(0, 0, 0, 0);
                    
                    switch (filter.dateFilter) {
                        case 'today':
                            if (dateValue.toDateString() !== today.toDateString()) includeRow = false;
                            break;
                        case 'tomorrow':
                            const tomorrow = new Date(today);
                            tomorrow.setDate(today.getDate() + 1);
                            if (dateValue.toDateString() !== tomorrow.toDateString()) includeRow = false;
                            break;
                        case 'yesterday':
                            const yesterday = new Date(today);
                            yesterday.setDate(today.getDate() - 1);
                            if (dateValue.toDateString() !== yesterday.toDateString()) includeRow = false;
                            break;
                        case 'thisWeek':
                            const startOfWeek = new Date(today);
                            const dayOfWeek = today.getDay();
                            const diffToMonday = dayOfWeek === 0 ? 6 : dayOfWeek - 1;
                            startOfWeek.setDate(today.getDate() - diffToMonday);
                            startOfWeek.setHours(0, 0, 0, 0);
                            const endOfWeek = new Date(startOfWeek);
                            endOfWeek.setDate(startOfWeek.getDate() + 6);
                            endOfWeek.setHours(23, 59, 59, 999);
                            if (dateValue < startOfWeek || dateValue > endOfWeek) includeRow = false;
                            break;
                        case 'lastWeek':
                            const lastWeekStart = new Date(today);
                            lastWeekStart.setDate(today.getDate() - today.getDay() - 6);
                            lastWeekStart.setHours(0, 0, 0, 0);
                            const lastWeekEnd = new Date(lastWeekStart);
                            lastWeekEnd.setDate(lastWeekStart.getDate() + 6);
                            lastWeekEnd.setHours(23, 59, 59, 999);
                            if (dateValue < lastWeekStart || dateValue > lastWeekEnd) includeRow = false;
                            break;
                        case 'nextWeek':
                            const nextWeekStart = new Date(today);
                            nextWeekStart.setDate(today.getDate() - today.getDay() + 7);
                            nextWeekStart.setHours(0, 0, 0, 0);
                            const nextWeekEnd = new Date(nextWeekStart);
                            nextWeekEnd.setDate(nextWeekStart.getDate() + 6);
                            nextWeekEnd.setHours(23, 59, 59, 999);
                            if (dateValue < nextWeekStart || dateValue > nextWeekEnd) includeRow = false;
                            break;
                        case 'thisMonth':
                            const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
                            const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);
                            endOfMonth.setHours(23, 59, 59, 999);
                            if (dateValue < startOfMonth || dateValue > endOfMonth) includeRow = false;
                            break;
                        case 'lastMonth':
                            const lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
                            const lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);
                            lastMonthEnd.setHours(23, 59, 59, 999);
                            if (dateValue < lastMonthStart || dateValue > lastMonthEnd) includeRow = false;
                            break;
                        case 'nextMonth':
                            const nextMonthStart = new Date(today.getFullYear(), today.getMonth() + 1, 1);
                            const nextMonthEnd = new Date(today.getFullYear(), today.getMonth() + 2, 0);
                            nextMonthEnd.setHours(23, 59, 59, 999);
                            if (dateValue < nextMonthStart || dateValue > nextMonthEnd) includeRow = false;
                            break;
                        case 'thisYear':
                            const startOfYear = new Date(today.getFullYear(), 0, 1);
                            const endOfYear = new Date(today.getFullYear(), 11, 31);
                            endOfYear.setHours(23, 59, 59, 999);
                            if (dateValue < startOfYear || dateValue > endOfYear) includeRow = false;
                            break;
                        case 'custom':
                            if (filter.customRange) {
                                if (dateValue < filter.customRange.start || dateValue > filter.customRange.end) {
                                    includeRow = false;
                                }
                            }
                            break;
                    }
                }
            } else {
                // Текстовый фильтр: массив отмеченных значений (чекбоксы как в Excel).
                // Пустой массив = "ничего не выбрано" = показать 0 строк.
                if (Array.isArray(filter.checkedValues)) {
                    if (filter.checkedValues.length === 0) {
                        includeRow = false;
                        break;
                    }
                    const cellValue = row[col];
                    const cellStr = cellValue === null || cellValue === undefined ? '' : String(cellValue).trim();
                    const checkedSet = new Set(filter.checkedValues.map(v => String(v).toLowerCase()));
                    if (!checkedSet.has(cellStr.toLowerCase())) {
                        includeRow = false;
                        break;
                    }
                }
            }
        }

        // Глобальный поиск: ячейка любого столбца должна содержать строку поиска
        if (includeRow && globalNeedle) {
            const found = row.some(cellValue => {
                if (cellValue === null || cellValue === undefined || cellValue === '') return false;
                return String(cellValue).toLowerCase().includes(globalNeedle);
            });
            if (!found) includeRow = false;
        }

        if (includeRow) {
            filteredData.push(row);
        }
    }

    // Сортировка (единая, не по каждому столбцу отдельно)
    if (nkitkSort && nkitkSort.order && nkitkSort.order !== 'none') {
        const isDate = isDateColumn(nkitkSort.col, nkitkTableData);
        filteredData = isDate
            ? sortByDate(filteredData, nkitkSort.col, nkitkSort.order)
            : sortByText(filteredData, nkitkSort.col, nkitkSort.order);
    }

    return filteredData;
}

// Открытие popover-фильтра (Excel-стиль: чекбоксы + поиск внутри).
// Для дат — блок быстрых фильтров и диапазона.
function openFilterPopover(anchorBtn, colIndex, colName) {
    const popover = document.getElementById('nkitk-filter-popover');
    if (!popover) return;

    // Повторный клик по той же кнопке — закрыть
    if (!popover.hidden && currentFilterCol === colIndex) {
        closeFilterPopover();
        return;
    }

    currentFilterCol = colIndex;
    const isDate = isDateColumn(colIndex, nkitkTableData);
    const existing = nkitkFilters[colIndex] || {};

    if (isDate) {
        renderDatePopover(popover, colIndex, colName, existing);
    } else {
        renderCheckboxPopover(popover, colIndex, colName, existing);
    }

    positionPopover(popover, anchorBtn);
    popover.hidden = false;
}

function positionPopover(popover, anchor) {
    popover.style.visibility = 'hidden';
    popover.hidden = false;
    const rect = anchor.getBoundingClientRect();
    const popW = popover.offsetWidth;
    const popH = popover.offsetHeight;
    const viewportW = window.innerWidth;
    const viewportH = window.innerHeight;

    let top = rect.bottom + window.scrollY + 4;
    let left = rect.left + window.scrollX;

    // Не вылезать за правый край
    if (left + popW > viewportW + window.scrollX - 8) {
        left = Math.max(8, viewportW + window.scrollX - popW - 8);
    }
    // Если не помещается снизу — показать сверху
    if (rect.bottom + popH + 8 > viewportH && rect.top - popH > 8) {
        top = rect.top + window.scrollY - popH - 4;
    }

    popover.style.top = top + 'px';
    popover.style.left = left + 'px';
    popover.style.visibility = '';
}

function closeFilterPopover() {
    const popover = document.getElementById('nkitk-filter-popover');
    if (popover) {
        popover.hidden = true;
        popover.innerHTML = '';
    }
    currentFilterCol = null;
}

function renderCheckboxPopover(popover, colIndex, colName, existing) {
    const uniqueValues = getUniqueValuesFromColumn(nkitkTableData, colIndex);
    const checked = new Set((existing.checkedValues || []).map(v => String(v).toLowerCase()));
    // По умолчанию — все значения отмечены (если фильтр не задан)
    const noFilter = !existing.checkedValues;

    const listItems = uniqueValues.map(v => {
        const isChecked = noFilter || checked.has(String(v).toLowerCase());
        const safeVal = escapeHtml(v);
        return `
            <label class="nkitk-popover__item">
                <input type="checkbox" class="nkitk-popover__checkbox" value="${safeVal}" ${isChecked ? 'checked' : ''}>
                <span class="nkitk-popover__text">${safeVal}</span>
            </label>
        `;
    }).join('');

    popover.innerHTML = `
        <div class="nkitk-popover__header">
            <div class="nkitk-popover__title">Фильтр: ${escapeHtml(colName)}</div>
            <button type="button" class="nkitk-popover__close" title="Закрыть">✖</button>
        </div>
        <input type="text" class="nkitk-popover__search" placeholder="Найти в списке..." autofocus>
        <div class="nkitk-popover__toolbar">
            <button type="button" class="nkitk-popover__action" data-act="all">Выбрать всё</button>
            <button type="button" class="nkitk-popover__action" data-act="none">Снять всё</button>
            ${uniqueValues.length === 0 ? '' : `<span class="nkitk-popover__count"></span>`}
        </div>
        <div class="nkitk-popover__list">
            ${uniqueValues.length > 0 ? listItems : '<div class="nkitk-popover__empty">Нет значений</div>'}
        </div>
        <div class="nkitk-popover__footer">
            <button type="button" class="nkitk-btn nkitk-btn--secondary nkitk-btn--small" data-act="reset">Сбросить фильтр столбца</button>
        </div>
    `;

    const searchInput = popover.querySelector('.nkitk-popover__search');
    const listEl = popover.querySelector('.nkitk-popover__list');
    const countEl = popover.querySelector('.nkitk-popover__count');

    function updateCount() {
        if (!countEl) return;
        const total = listEl.querySelectorAll('.nkitk-popover__item').length;
        const checkedCount = listEl.querySelectorAll('.nkitk-popover__checkbox:checked').length;
        countEl.textContent = `${checkedCount}/${total}`;
    }
    updateCount();

    // Поиск внутри списка
    searchInput.addEventListener('input', () => {
        const needle = searchInput.value.trim().toLowerCase();
        listEl.querySelectorAll('.nkitk-popover__item').forEach(item => {
            const text = item.querySelector('.nkitk-popover__text').textContent.toLowerCase();
            item.style.display = text.includes(needle) ? '' : 'none';
        });
    });

    // Применение изменений чекбоксов — мгновенно
    function applyCheckboxState() {
        const visibleItems = Array.from(listEl.querySelectorAll('.nkitk-popover__item'))
            .filter(it => it.style.display !== 'none');
        const allBoxes = Array.from(listEl.querySelectorAll('.nkitk-popover__checkbox'));
        const checkedBoxes = allBoxes.filter(cb => cb.checked);

        // Если отмечены все — фильтр считается снятым
        if (checkedBoxes.length === allBoxes.length) {
            delete nkitkFilters[colIndex];
        } else {
            nkitkFilters[colIndex] = {
                type: 'text',
                checkedValues: checkedBoxes.map(cb => cb.value)
            };
        }
        updateCount();
        renderNKiTKTable();
    }

    listEl.addEventListener('change', (e) => {
        if (e.target.classList.contains('nkitk-popover__checkbox')) applyCheckboxState();
    });

    popover.querySelectorAll('.nkitk-popover__action').forEach(btn => {
        btn.addEventListener('click', () => {
            const act = btn.dataset.act;
            listEl.querySelectorAll('.nkitk-popover__item').forEach(item => {
                if (item.style.display === 'none') return;
                const cb = item.querySelector('.nkitk-popover__checkbox');
                cb.checked = (act === 'all');
            });
            applyCheckboxState();
        });
    });

    popover.querySelector('[data-act="reset"]').addEventListener('click', () => {
        delete nkitkFilters[colIndex];
        closeFilterPopover();
        renderNKiTKTable();
    });

    popover.querySelector('.nkitk-popover__close').addEventListener('click', closeFilterPopover);
}

function renderDatePopover(popover, colIndex, colName, existing) {
    popover.innerHTML = `
        <div class="nkitk-popover__header">
            <div class="nkitk-popover__title">Фильтр по дате: ${escapeHtml(colName)}</div>
            <button type="button" class="nkitk-popover__close" title="Закрыть">✖</button>
        </div>
        <div class="filter-form-group">
            <label class="filter-form-label">Быстрые фильтры:</label>
            <select id="date-quick-filter" class="filter-form-select">
                <option value="all">-- Все даты --</option>
                <option value="today">📅 Сегодня</option>
                <option value="tomorrow">📅 Завтра</option>
                <option value="yesterday">📅 Вчера</option>
                <option value="thisWeek">📅 Текущая неделя</option>
                <option value="lastWeek">📅 Прошлая неделя</option>
                <option value="nextWeek">📅 Следующая неделя</option>
                <option value="thisMonth">📅 Текущий месяц</option>
                <option value="lastMonth">📅 Прошлый месяц</option>
                <option value="nextMonth">📅 Следующий месяц</option>
                <option value="thisYear">📅 Текущий год</option>
            </select>
        </div>
        <div class="filter-form-group">
            <label class="filter-form-label">Произвольный диапазон:</label>
            <div class="filter-form-date-range">
                <input type="date" id="date-start" class="filter-form-input" placeholder="Дата от">
                <input type="date" id="date-end" class="filter-form-input" placeholder="Дата до">
            </div>
            <button id="apply-custom-range" class="nkitk-btn nkitk-btn--block">Применить диапазон</button>
        </div>
        <div class="nkitk-popover__footer">
            <button type="button" class="nkitk-btn nkitk-btn--secondary nkitk-btn--small" data-act="reset">Сбросить фильтр столбца</button>
        </div>
    `;

    const quickFilter = popover.querySelector('#date-quick-filter');
    if (existing.dateFilter && existing.dateFilter !== 'custom') quickFilter.value = existing.dateFilter;

    if (existing.customRange) {
        const sI = popover.querySelector('#date-start');
        const eI = popover.querySelector('#date-end');
        if (sI && existing.customRange.start) sI.value = existing.customRange.start.toISOString().split('T')[0];
        if (eI && existing.customRange.end) eI.value = existing.customRange.end.toISOString().split('T')[0];
    }

    quickFilter.addEventListener('change', () => {
        const v = quickFilter.value;
        if (v === 'all') delete nkitkFilters[colIndex];
        else nkitkFilters[colIndex] = { type: 'date', dateFilter: v };
        renderNKiTKTable();
    });

    popover.querySelector('#apply-custom-range').addEventListener('click', () => {
        const sI = popover.querySelector('#date-start');
        const eI = popover.querySelector('#date-end');
        if (!sI.value || !eI.value) {
            alert('Выберите обе даты');
            return;
        }
        nkitkFilters[colIndex] = {
            type: 'date',
            dateFilter: 'custom',
            customRange: { start: new Date(sI.value), end: new Date(eI.value) }
        };
        renderNKiTKTable();
    });

    popover.querySelector('[data-act="reset"]').addEventListener('click', () => {
        delete nkitkFilters[colIndex];
        closeFilterPopover();
        renderNKiTKTable();
    });

    popover.querySelector('.nkitk-popover__close').addEventListener('click', closeFilterPopover);
}

// Закрытие popover по клику вне него или Escape
document.addEventListener('click', (e) => {
    const popover = document.getElementById('nkitk-filter-popover');
    if (!popover || popover.hidden) return;
    if (popover.contains(e.target)) return;
    if (e.target.closest('.nkitk-filter-btn')) return;
    closeFilterPopover();
});
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeFilterPopover();
});

window.clearAllFilters = function() {
    nkitkFilters = {};
    nkitkGlobalSearch = '';
    nkitkSort = null;
    closeFilterPopover();
    renderNKiTKTable();
};

function updateActiveFiltersDisplay() {
    const displayDiv = document.getElementById('active-filters-display');
    if (!displayDiv) return;

    const chips = [];
    const headers = nkitkTableData[0] || [];

    // Глобальный поиск
    if (nkitkGlobalSearch && nkitkGlobalSearch.trim() !== '') {
        chips.push({
            label: `🔍 поиск: "${nkitkGlobalSearch.trim()}"`,
            onRemove: 'window.clearGlobalSearch()'
        });
    }

    // Сортировка
    if (nkitkSort && nkitkSort.order) {
        const colName = headers[nkitkSort.col] || `Колонка ${String.fromCharCode(65 + nkitkSort.col)}`;
        chips.push({
            label: `↕ ${colName}: ${nkitkSort.order === 'asc' ? 'А→Я' : 'Я→А'}`,
            onRemove: 'window.clearSort()'
        });
    }

    // Фильтры колонок
    Object.entries(nkitkFilters).forEach(([col, filter]) => {
        const colName = headers[parseInt(col)] || `Колонка ${String.fromCharCode(65 + parseInt(col))}`;
        let filterText = '';

        if (filter.type === 'date') {
            if (filter.dateFilter === 'custom' && filter.customRange) {
                filterText = `${filter.customRange.start.toLocaleDateString()}–${filter.customRange.end.toLocaleDateString()}`;
            } else {
                const filterNames = {
                    'today': 'Сегодня', 'tomorrow': 'Завтра', 'yesterday': 'Вчера',
                    'thisWeek': 'Тек. неделя', 'lastWeek': 'Прошл. неделя', 'nextWeek': 'След. неделя',
                    'thisMonth': 'Тек. месяц', 'lastMonth': 'Прошл. месяц', 'nextMonth': 'След. месяц',
                    'thisYear': 'Тек. год'
                };
                filterText = filterNames[filter.dateFilter] || filter.dateFilter;
            }
        } else if (Array.isArray(filter.checkedValues)) {
            if (filter.checkedValues.length <= 2) {
                filterText = filter.checkedValues.map(v => `"${v}"`).join(', ');
            } else {
                filterText = `выбрано ${filter.checkedValues.length}`;
            }
        }

        if (filterText) {
            chips.push({
                label: `${colName}: ${filterText}`,
                onRemove: `window.removeFilter(${col})`
            });
        }
    });

    if (chips.length === 0) {
        displayDiv.innerHTML = '';
        displayDiv.hidden = true;
        return;
    }

    displayDiv.hidden = false;
    displayDiv.innerHTML = chips.map(c => `
        <span class="active-filter-chip">
            ${escapeHtml(c.label)}
            <button onclick="${c.onRemove}" class="active-filter-chip__remove" title="Убрать">✖</button>
        </span>
    `).join('');
}

window.clearGlobalSearch = function() {
    nkitkGlobalSearch = '';
    renderNKiTKTable();
};

window.clearSort = function() {
    nkitkSort = null;
    renderNKiTKTable();
};

window.removeFilter = function(colIndex) {
    delete nkitkFilters[colIndex];
    renderNKiTKTable();
};

// ============================================================
// ЗАГРУЗКА ФАЙЛА
// ============================================================

async function loadExcel() {
    const refreshBtn = document.getElementById('refresh-btn');

    try {
        if (typeof XLSX === 'undefined') {
            throw new Error('Библиотека SheetJS (XLSX) не загружена. Проверьте подключение к интернету.');
        }
        if (typeof Chart === 'undefined') {
            throw new Error('Библиотека Chart.js не загружена. Проверьте подключение к интернету.');
        }

        if (refreshBtn) {
            refreshBtn.disabled = true;
            refreshBtn.classList.add('navbar__refresh--spinning');
        }
        setStatus('Загрузка данных...', 'loading');

        // Таймаут 30 секунд — если сервер не ответил, прерываем запрос
        const FETCH_TIMEOUT_MS = 30000;
        const abortController = new AbortController();
        const timeoutId = setTimeout(() => abortController.abort(), FETCH_TIMEOUT_MS);

        let response;
        try {
            response = await fetch(EXCEL_FILE_PATH + '?v=' + Date.now(), {
                cache: 'no-store',
                signal: abortController.signal
            });
        } catch (fetchError) {
            if (fetchError.name === 'AbortError') {
                throw new Error(`Сервер не ответил за ${FETCH_TIMEOUT_MS / 1000} секунд. Проверьте сеть.`);
            }
            throw fetchError;
        } finally {
            clearTimeout(timeoutId);
        }

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const lastModified = response.headers.get('last-modified');
        if (lastModified) {
            fileLastModified = new Date(lastModified);
        }
        
        const buffer = await response.arrayBuffer();
        
        workbook = XLSX.read(buffer, {
            type: 'array',
            raw: true,
            cellDates: false
        });
        resetWorkbookDataCache();
        techBlockStatusRulesCache = null;
        
        log('Доступные листы:', workbook.SheetNames);
        
        const defaultSheet = workbook.SheetNames[0];
        await loadSheetData(defaultSheet);
        
        setStatus('Данные загружены', 'success');
        
    } catch (error) {
        console.error('Ошибка загрузки:', error);
        setStatus(`Ошибка загрузки: ${error.message || 'неизвестная ошибка'}`, 'error');
    } finally {
        if (refreshBtn) {
            refreshBtn.disabled = false;
            refreshBtn.classList.remove('navbar__refresh--spinning');
        }
    }
}

function loadSheetData(sheetName) {
    return new Promise((resolve) => {
        if (!workbook) {
            console.warn('Workbook не загружен');
            resolve();
            return;
        }

        if (sheetName === 'Расчеты') {
            chartManager.destroyAll();
            currentSheet = sheetName;
            switchToCalculationsMode();
            activateNavButton(sheetName);
            updateFileInfo();
            resolve();
            return;
        }

        if (sheetName === 'ПП') {
            chartManager.destroyAll();
            currentSheet = sheetName;

            worksheetCache = getCachedWorksheetCache('ТехБлок');

            switchToProductionProgramMode();
            activateNavButton(sheetName);
            updateFileInfo();
            resolve();
            return;
        }
        
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            console.error(`Лист ${sheetName} не найден`);
            resolve();
            return;
        }
        
        log(`Загружаем лист: ${sheetName}`);
        
        worksheetCache = getCachedWorksheetCache(sheetName);
        currentSheet = sheetName;
        
        if (sheetName === 'НКиТК') {
            chartManager.destroyAll();
            switchToNKiTKMode();
        } else {
            chartManager.destroyAll();
            switchToMainMode();
        }
        
        activateNavButton(sheetName);
        updateFileInfo();
        
        resolve();
    });
}

function switchToNKiTKMode() {
    isNKiTKMode = true;
    // Фильтры намеренно НЕ сбрасываем — сохраняем между переключениями вкладок.
    // Сброс только через кнопку «Сбросить все фильтры» (clearAllFilters).

    const mainContent = document.getElementById('main-content');
    if (mainContent) mainContent.style.display = 'none';
    
    const nkitkContent = document.getElementById('nkitk-content');
    if (nkitkContent) {
        nkitkContent.style.display = 'block';
        const title = nkitkContent.querySelector('.center-title');
        if (title) {
            title.textContent = 'Проекты центров для НК и ТК';
        }
    }

    const calculationsContent = document.getElementById('calculations-content');
    if (calculationsContent) calculationsContent.style.display = 'none';

    const productionProgramContent = document.getElementById('production-program-content');
    if (productionProgramContent) productionProgramContent.style.display = 'none';
    
    renderNKiTKTable();
}

function switchToCalculationsMode() {
    isNKiTKMode = false;

    const mainContent = document.getElementById('main-content');
    if (mainContent) mainContent.style.display = 'none';

    const nkitkContent = document.getElementById('nkitk-content');
    if (nkitkContent) nkitkContent.style.display = 'none';

    const calculationsContent = document.getElementById('calculations-content');
    if (calculationsContent) calculationsContent.style.display = 'block';

    const productionProgramContent = document.getElementById('production-program-content');
    if (productionProgramContent) productionProgramContent.style.display = 'none';

    renderCalculationsAudit();
}

function switchToProductionProgramMode() {
    isNKiTKMode = false;

    const mainContent = document.getElementById('main-content');
    if (mainContent) mainContent.style.display = 'none';

    const nkitkContent = document.getElementById('nkitk-content');
    if (nkitkContent) nkitkContent.style.display = 'none';

    const calculationsContent = document.getElementById('calculations-content');
    if (calculationsContent) calculationsContent.style.display = 'none';

    const productionProgramContent = document.getElementById('production-program-content');
    if (productionProgramContent) productionProgramContent.style.display = 'block';

    renderProductionProgramPlanFact();
}

function switchToMainMode() {
    isNKiTKMode = false;
    
    const mainContent = document.getElementById('main-content');
    if (mainContent) mainContent.style.display = 'block';
    
    const nkitkContent = document.getElementById('nkitk-content');
    if (nkitkContent) nkitkContent.style.display = 'none';

    const calculationsContent = document.getElementById('calculations-content');
    if (calculationsContent) calculationsContent.style.display = 'none';

    const productionProgramContent = document.getElementById('production-program-content');
    if (productionProgramContent) productionProgramContent.style.display = 'none';
    
    renderAll();
}

// ============================================================
// РЕНДЕРИНГ ОСНОВНОГО ДАШБОРДА
// ============================================================

function renderAll() {
    if (!worksheetCache) {
        console.warn('Нет данных для рендеринга');
        return;
    }
    
    log('Рендерим все компоненты...');
    
    try {
        updateWorkloadBlockTitle();
        updateHeaderInfo();
        updateDashboardDate();
        updateNotification();
        updateTargetIndicator();
        updateIndicatorsLegend();
        updateEmployeeCounts();
        updateAbsences();
        updateLoadingAttention();
        renderChartCards();
        renderCenterMonthlyLoadSection();
        renderQuarterlyTrendSection();
        updateProjects();

        log('Рендеринг завершен');
    } catch (error) {
        console.error('Ошибка при рендеринге:', error);
    }
}

function updateWorkloadBlockTitle() {
    const title = document.querySelector('#main-content .workload-header-title-row h3');
    if (!title) return;

    title.textContent = currentSheet === 'ТехБлок' ? 'Загрузка блока' : 'Загрузка центра';
}

function updateDashboardDate() {
    try {
        const wrapper = document.getElementById('dashboard-updated-at');
        const valueEl = document.getElementById('dashboard-updated-at-value');
        const timeEl = document.getElementById('dashboard-updated-at-time');
        if (!wrapper || !valueEl) return;

        const date = fileLastModified || new Date();

        if (timeEl) {
            timeEl.textContent = date.toLocaleTimeString('ru-RU', {
                hour: '2-digit',
                minute: '2-digit'
            });
        }

        valueEl.textContent = date.toLocaleDateString('ru-RU', {
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        });

        wrapper.style.display = 'inline-flex';
    } catch (error) {
        console.warn('Ошибка в updateDashboardDate:', error);
    }
}

function updateHeaderInfo() {
    try {
        const centerName = getCellValue(EXCEL_CONFIG.header.centerName, worksheetCache);
        const titleEl = document.querySelector('#main-content .center-title');
        if (titleEl && isValid(centerName)) {
            titleEl.textContent = String(centerName).trim();
        }
        
        const manager = getCellValue(EXCEL_CONFIG.header.manager, worksheetCache);
        const managerEl = document.querySelector('#main-content .header-info h6:first-child');
        if (managerEl) {
            if (isValid(manager)) {
                managerEl.textContent = `Руководитель центра - ${String(manager).trim()}`;
                managerEl.style.display = '';
            } else {
                managerEl.textContent = '';
                managerEl.style.display = 'none';
            }
        }
        
        const admin = getCellValue(EXCEL_CONFIG.header.admin, worksheetCache);
        const adminEl = document.querySelector('#main-content .header-info h6:last-child');
        if (adminEl) {
            if (isValid(admin)) {
                adminEl.textContent = `Администратор центра - ${String(admin).trim()}`;
                adminEl.style.display = '';
            } else {
                adminEl.textContent = '';
                adminEl.style.display = 'none';
            }
        }
    } catch (error) {
        console.warn('Ошибка в updateHeaderInfo:', error);
    }
}

function updateNotification() {
    try {
        const notificationData = getNotificationData();
        const notificationBlock = document.getElementById('notification-block');
        const notificationContent = document.getElementById('notification-content');

        if (!notificationBlock || !notificationContent) return;

        if (notificationData.entries.length > 0) {
            const rowsHtml = notificationData.entries.map(entry => `
                <div class="notification-table__row">
                    <span class="notification-table__label">${escapeHtml(entry.label)}</span>
                    <span class="notification-table__value">${escapeHtml(entry.valueText)}</span>
                </div>
            `).join('');

            notificationContent.innerHTML = `
                <div class="notification-summary">
                    <span class="notification-summary__label">Контрольные показатели по проектам</span>
                    <span class="notification-summary__badge">${notificationData.attentionCount}</span>
                </div>
                <div class="notification-table">${rowsHtml}</div>
            `;
            notificationBlock.style.display = 'block';

            const h3 = notificationBlock.querySelector('h3');
            if (h3) {
                h3.textContent = 'Уведомления по проектам';
            }
        } else if (notificationData.fallbackText) {
            notificationContent.innerHTML = `<div class="notification-note">${escapeHtml(notificationData.fallbackText)}</div>`;
            notificationBlock.style.display = 'block';

            const h3 = notificationBlock.querySelector('h3');
            if (h3) {
                h3.textContent = 'Уведомления по проектам';
            }
        } else {
            notificationBlock.style.display = 'none';
        }
    } catch (error) {
        console.warn('Ошибка в updateNotification:', error);
    }
}

function updateTargetIndicator() {
    try {
        const block = document.getElementById('target-indicator-block');
        const content = document.getElementById('target-indicator-content');
        if (!block || !content) return;

        const data = getTargetIndicatorData();
        if (!data || data.hasSource === false) {
            block.style.display = 'none';
            content.innerHTML = '';
            return;
        }

        content.innerHTML = renderTargetIndicatorContent(data);
        block.style.display = 'block';
        attachTargetIndicatorPeriodHandler();
    } catch (error) {
        console.warn('Ошибка в updateTargetIndicator:', error);
    }
}

const TARGET_INDICATOR_YEAR = 2026;
const TARGET_INDICATOR_EXCLUDED_CENTERS = new Set(['УВП']);
const TARGET_INDICATOR_NO_KD_CENTERS = new Set(['ЛОИ']);
const TARGET_INDICATOR_TECHBLOCK_OTHER_CENTERS_KEY = 'Остальные центры';
const TARGET_INDICATOR_TECHBLOCK_OTHER_CENTERS_LABEL = 'Центры КБ';
const TARGET_INDICATOR_ALLOWED_PROJECT_TYPES = new Set([
    'коммерческий',
    'коммерческий проект',
    'инвестиционный',
    'инвестиционный проект'
]);

function getTargetIndicatorAvailablePeriods(referenceDate = new Date()) {
    const periods = [
        {
            key: '2026-H1',
            year: TARGET_INDICATOR_YEAR,
            label: 'I полугодие 2026',
            shortLabel: 'I полугодие 2026',
            start: new Date(TARGET_INDICATOR_YEAR, 0, 1),
            end: new Date(TARGET_INDICATOR_YEAR, 5, 30, 23, 59, 59, 999)
        }
    ];

    if (referenceDate >= new Date(TARGET_INDICATOR_YEAR, 6, 1)) {
        periods.push({
            key: '2026-H2',
            year: TARGET_INDICATOR_YEAR,
            label: 'II полугодие 2026',
            shortLabel: 'II полугодие 2026',
            start: new Date(TARGET_INDICATOR_YEAR, 6, 1),
            end: new Date(TARGET_INDICATOR_YEAR, 11, 31, 23, 59, 59, 999)
        });
    }

    return periods;
}

function getTargetIndicatorPeriod(periodKey = targetIndicatorPeriodKey) {
    const periods = getTargetIndicatorAvailablePeriods();
    return periods.find(period => period.key === periodKey) || periods[0];
}

function isDateInTargetIndicatorPeriod(date, period) {
    return date instanceof Date
        && !isNaN(date.getTime())
        && period
        && date >= period.start
        && date <= period.end;
}

function getTargetIndicatorData(periodKey = targetIndicatorPeriodKey) {
    const periods = getTargetIndicatorAvailablePeriods();
    const period = getTargetIndicatorPeriod(periodKey);
    targetIndicatorPeriodKey = period.key;
    const currentCenter = normalizeProjectCenterLabel(currentSheet);

    if (currentCenter && TARGET_INDICATOR_EXCLUDED_CENTERS.has(currentCenter)) {
        return { period, periods, rows: [], hasSource: false };
    }

    if (!workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) {
        return { period, periods, rows: [], hasSource: false };
    }

    const projects = getTargetIndicatorProjectRecords();
    const centerMap = new Map();

    projects.forEach(project => {
        if (!isTargetIndicatorProjectStartedInPeriod(project, period)) return;

        const group = getTargetIndicatorDisplayGroup(project.center);
        if (!group) return;

        if (!centerMap.has(group.key)) {
            centerMap.set(group.key, {
                center: group.key,
                centerName: group.label,
                total: 0,
                completed: 0,
                onTime: 0,
                inHours: 0,
                hoursOverrun: 0
            });
        }

        const center = centerMap.get(group.key);
        const isCompleted = isTargetIndicatorProjectCompletedInPeriod(project, period);
        const isOnTime = isCompleted
            && project.planEnd instanceof Date
            && !isNaN(project.planEnd.getTime())
            && project.factEnd <= project.planEnd;
        const isInHours = isCompleted
            && isTargetIndicatorProjectInHours(project);

        center.total += 1;
        if (isCompleted) center.completed += 1;
        if (isOnTime) center.onTime += 1;
        if (isInHours) center.inHours += 1;
        if (isCompleted) center.hoursOverrun += getTargetIndicatorProjectHoursOverrun(project);
    });

    let resultRows = getOrderedTargetIndicatorRows(Array.from(centerMap.values()))
        .map(row => ({
            ...row,
            onTimePercent: calculateTargetIndicatorPercent(row.onTime, row.completed),
            inHoursPercent: calculateTargetIndicatorPercent(row.inHours, row.completed)
        }))
        .filter(row => row.total > 0);

    if (currentCenter && currentSheet !== 'ТехБлок') {
        resultRows = resultRows.filter(row => row.center === currentCenter);
    }

    return {
        year: period.year,
        period,
        periods,
        rows: resultRows,
        mode: currentSheet === 'ТехБлок' ? 'all' : 'center',
        hasSource: true
    };
}

function getTargetIndicatorProjectRecords() {
    if (!workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) {
        return [];
    }

    const rows = getWorksheetRowsByHeaders('Проекты_ТехБлок');
    const projectMap = new Map();

    rows.forEach((row, index) => {
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const id = normalizeProjectId(getRowValue(row, ['ID']));
        if (!center || !id) return;

        const projectType = getCleanTextValue(getRowValue(row, ['Вид проекта']));
        if (!isTargetIndicatorProjectTypeIncluded(projectType)) return;

        const key = getTargetIndicatorProjectKey(center, id);
        if (!projectMap.has(key)) {
            projectMap.set(key, {
                key,
                center,
                id,
                projectType: '',
                projectTypes: [],
                statuses: [],
                planEnd: null,
                planHours: 0,
                actualHours: 0,
                hasPlanHours: false,
                hasReadyKd: false,
                hasArchivedStatus: false,
                factStart: null,
                factEnd: null,
                sourceIndex: index
            });
        }

        const project = projectMap.get(key);
        if (!project.projectType && projectType) project.projectType = projectType;
        if (projectType) project.projectTypes.push(projectType);

        project.statuses.push(getRowValue(row, ['Статус']));
        if (isTargetIndicatorKdReady(getRowValue(row, [
            'Статус КД',
            'Статус РКД',
            'КД',
            'Столбец 17'
        ]))) {
            project.hasReadyKd = true;
        }
        if (isTargetIndicatorArchivedStatusValue(getRowValue(row, ['Статус']))) {
            project.hasArchivedStatus = true;
        }

        project.planEnd = getLaterDate(project.planEnd, parseDateValue(getRowValue(row, [
            'План дата завершения этапа по графику проекта',
            'Дата завершения работ',
            'Дата завершения работ по договору',
            'Плановая дата окончания'
        ])));

        const planHours = parseNumericValue(getRowValue(row, [
            'Количество ПЛАНовых часов',
            'Плановые часы'
        ]));
        if (planHours !== null) {
            project.planHours += planHours;
            project.hasPlanHours = true;
        }

        const actualHours = parseNumericValue(getRowValue(row, [
            'Количество ФАКТических часов',
            'Фактические часы'
        ]));
        if (actualHours !== null) {
            project.actualHours += actualHours;
        }

        project.factStart = getEarlierDate(project.factStart, parseDateValue(getRowValue(row, [
            'ФАКТ начало',
            'Фактическая дата начала'
        ])));
        project.factEnd = getLaterDate(project.factEnd, parseDateValue(getRowValue(row, [
            'ФАКТ окончание',
            'Фактическая дата окончания'
        ])));
    });

    return Array.from(projectMap.values());
}

function getTargetIndicatorProjectKey(center, id) {
    const normalizedCenter = normalizeProjectCenterLabel(center);
    const normalizedId = normalizeProjectId(id);
    return normalizedCenter && normalizedId ? `${normalizedCenter}||${normalizedId}` : '';
}

function getTargetIndicatorDisplayGroup(center) {
    const normalizedCenter = normalizeProjectCenterLabel(center);
    if (!normalizedCenter || TARGET_INDICATOR_EXCLUDED_CENTERS.has(normalizedCenter)) return null;

    if (currentSheet === 'ТехБлок') {
        if (normalizedCenter === 'ЛОИ') {
            return { key: 'ЛОИ', label: 'ЛОИ' };
        }

        return {
            key: TARGET_INDICATOR_TECHBLOCK_OTHER_CENTERS_KEY,
            label: TARGET_INDICATOR_TECHBLOCK_OTHER_CENTERS_LABEL
        };
    }

    return {
        key: normalizedCenter,
        label: getProjectStageFullLabel(normalizedCenter)
    };
}

function isTargetIndicatorProjectStartedInYear(project, year) {
    const period = {
        start: new Date(year, 0, 1),
        end: new Date(year, 11, 31, 23, 59, 59, 999)
    };

    return isTargetIndicatorProjectStartedInPeriod(project, period);
}

function isTargetIndicatorProjectStartedInPeriod(project, period) {
    return Boolean(project)
        && !TARGET_INDICATOR_EXCLUDED_CENTERS.has(project.center)
        && isTargetIndicatorProjectTypeIncluded(project.projectTypes)
        && isTargetIndicatorProjectEligibleByKdOrStatus(project)
        && isDateInTargetIndicatorPeriod(project.factStart, period);
}

function isTargetIndicatorProjectCompletedInYear(project, year) {
    const period = {
        start: new Date(year, 0, 1),
        end: new Date(year, 11, 31, 23, 59, 59, 999)
    };

    return isTargetIndicatorProjectCompletedInPeriod(project, period);
}

function isTargetIndicatorProjectCompletedInPeriod(project, period) {
    return isTargetIndicatorProjectStartedInPeriod(project, period)
        && isTargetIndicatorProjectCompletedStatus(project)
        && isDateInTargetIndicatorPeriod(project.factEnd, period);
}

function getTargetIndicatorProjectKeySet(periodKey = targetIndicatorPeriodKey, options = {}) {
    const completedOnly = options.completedOnly === true;
    const period = getTargetIndicatorPeriod(periodKey);
    const keys = new Set();

    getTargetIndicatorProjectRecords().forEach(project => {
        const shouldInclude = completedOnly
            ? isTargetIndicatorProjectCompletedInPeriod(project, period)
            : isTargetIndicatorProjectStartedInPeriod(project, period);

        if (shouldInclude && project.key) keys.add(project.key);
    });

    return keys;
}

function getOrderedTargetIndicatorRows(rows) {
    const order = [TARGET_INDICATOR_TECHBLOCK_OTHER_CENTERS_KEY, 'ЛОИ', 'ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ'];

    return rows.sort((a, b) => {
        const aIndex = order.indexOf(a.center);
        const bIndex = order.indexOf(b.center);
        const aRank = aIndex === -1 ? order.length : aIndex;
        const bRank = bIndex === -1 ? order.length : bIndex;
        if (aRank !== bRank) return aRank - bRank;
        return (a.centerName || a.center).localeCompare(b.centerName || b.center, 'ru');
    });
}

function calculateTargetIndicatorPercent(value, total) {
    return total > 0 ? Math.round((value / total) * 100) : null;
}

function isTargetIndicatorProjectInHours(project) {
    if (!project || !project.hasPlanHours) return false;
    const diffHours = (project.planHours || 0) - (project.actualHours || 0);
    return diffHours >= -0.000001;
}

function getTargetIndicatorProjectHoursOverrun(project) {
    if (!project || !project.hasPlanHours) return 0;
    const overrun = (project.actualHours || 0) - (project.planHours || 0);
    return Number.isFinite(overrun) && overrun > 0 ? overrun : 0;
}

function isTargetIndicatorKdReady(value) {
    return normalizeLookupText(value) === 'готово';
}

function isTargetIndicatorProjectEligibleByKdOrStatus(project) {
    if (!project) return false;

    if (TARGET_INDICATOR_NO_KD_CENTERS.has(project.center)) {
        return isTargetIndicatorProjectCompleted(project.statuses);
    }

    return project.hasReadyKd === true || project.hasArchivedStatus === true;
}

function isTargetIndicatorProjectCompleted(statuses) {
    const normalizedStatuses = statuses.map(status => normalizeProjectStage(status));

    return normalizedStatuses.length > 0
        && normalizedStatuses.every(status => status.type === 'done');
}

function isTargetIndicatorArchivedStatusValue(value) {
    return normalizeLookupText(value) === 'архив';
}

function isTargetIndicatorProjectCompletedStatus(project) {
    if (!project) return false;
    if (isTargetIndicatorProjectCompleted(project.statuses)) return true;
    return !TARGET_INDICATOR_NO_KD_CENTERS.has(project.center) && project.hasArchivedStatus === true;
}

function isTargetIndicatorProjectTypeIncluded(projectTypes) {
    const values = Array.isArray(projectTypes) ? projectTypes : [projectTypes];

    return values.some(value => {
        const normalized = normalizeLookupText(value);
        return TARGET_INDICATOR_ALLOWED_PROJECT_TYPES.has(normalized);
    });
}

function renderTargetIndicatorContent(data) {
    const taskIds = getTargetIndicatorTaskIdRows(data.period.key, data.mode);
    const periodToggleHtml = renderTargetIndicatorPeriodToggle(data);
    const periodLabel = data.period?.label || String(data.year);

    if (!data.rows || data.rows.length === 0) {
        return `
            ${periodToggleHtml}
            <p class="target-indicator-subtitle">${escapeHtml(periodLabel)}: нет завершенных коммерческих и инвестиционных задач для выбранного среза. Для ЛОИ вместо статуса КД используется статус проекта «Завершена».</p>
            ${renderTargetIndicatorTaskIds(taskIds, data.mode)}
        `;
    }

    if (data.mode === 'all') {
        const rowsHtml = data.rows.map(row => `
            <div class="target-indicator-group-row">
                <div class="target-indicator-group-row__title">${escapeHtml(row.centerName || row.center)}</div>
                <div class="target-indicator-metrics target-indicator-metrics--grouped">
                    ${renderTargetIndicatorMetric('В срок', row.onTimePercent, row.onTime, row.completed)}
                    ${renderTargetIndicatorMetric('В норме часов', row.inHoursPercent, row.inHours, row.completed, {
                        extraNote: `превышение: ${formatHoursValue(row.hoursOverrun || 0)} ч`
                    })}
                </div>
            </div>
        `).join('');

        return `
            ${periodToggleHtml}
            <p class="target-indicator-subtitle">${escapeHtml(periodLabel)}: Центры КБ считаются суммарно без ЛОИ и УВП (учитывается статус КД «Готово» или статус проекта «Архив»), ЛОИ считается отдельно. Для ЛОИ вместо статуса КД используется статус проекта «Завершена». Доля в срок и по часам считается от завершенных задач выбранного полугодия.</p>
            <div class="target-indicator-groups">
                ${rowsHtml}
            </div>
            ${renderTargetIndicatorTaskIds(taskIds, data.mode)}
        `;
    }

    const row = data.rows[0];
    if (!row) return '';

    return `
        ${periodToggleHtml}
        <p class="target-indicator-subtitle">${escapeHtml(periodLabel)}: коммерческие и инвестиционные проекты центра со статусом КД «Готово» или статусом проекта «Архив»; для ЛОИ вместо статуса КД используется статус проекта «Завершена».</p>
        <div class="target-indicator-summary">
            <div>
                <span class="target-indicator-summary__label">Начато</span>
                <strong>${escapeHtml(String(row.total))}</strong>
            </div>
            <div>
                <span class="target-indicator-summary__label">Завершено</span>
                <strong>${escapeHtml(String(row.completed))}</strong>
            </div>
        </div>
        <div class="target-indicator-metrics">
            ${renderTargetIndicatorMetric('В срок', row.onTimePercent, row.onTime, row.completed)}
            ${renderTargetIndicatorMetric('В норме часов', row.inHoursPercent, row.inHours, row.completed, {
                extraNote: `превышение: ${formatHoursValue(row.hoursOverrun || 0)} ч`
            })}
        </div>
        ${renderTargetIndicatorTaskIds(taskIds, data.mode)}
    `;
}

function renderTargetIndicatorPeriodToggle(data) {
    const periods = Array.isArray(data.periods) ? data.periods : [];
    if (periods.length === 0) return '';

    return `
        <div class="target-indicator-period-toggle projects-segment-toggle" role="group" aria-label="Срез целевого показателя">
            ${periods.map(period => `
                <button type="button"
                        class="projects-segment-toggle__button${period.key === data.period.key ? ' is-active' : ''}"
                        data-target-indicator-period="${escapeHtml(period.key)}"
                        aria-pressed="${period.key === data.period.key ? 'true' : 'false'}">
                    ${escapeHtml(period.shortLabel)}
                </button>
            `).join('')}
        </div>
    `;
}

function attachTargetIndicatorPeriodHandler() {
    document.querySelectorAll('[data-target-indicator-period]').forEach(button => {
        button.addEventListener('click', () => {
            const nextPeriodKey = button.dataset.targetIndicatorPeriod;
            if (!nextPeriodKey || targetIndicatorPeriodKey === nextPeriodKey) return;

            targetIndicatorPeriodKey = nextPeriodKey;
            updateTargetIndicator();
            updateProjects();
        });
    });
}

function getTargetIndicatorTaskIdRows(periodKey, mode) {
    const currentCenter = normalizeProjectCenterLabel(currentSheet);
    const period = getTargetIndicatorPeriod(periodKey);

    return getTargetIndicatorProjectRecords()
        .filter(project => isTargetIndicatorProjectCompletedInPeriod(project, period))
        .filter(project => {
            if (mode === 'all') return Boolean(getTargetIndicatorDisplayGroup(project.center));
            return !currentCenter || project.center === currentCenter;
        })
        .sort((a, b) => {
            const centerOrder = getOrderedTargetIndicatorRows([
                { center: a.center, centerName: getProjectStageFullLabel(a.center) },
                { center: b.center, centerName: getProjectStageFullLabel(b.center) }
            ]);
            if (centerOrder[0]?.center !== centerOrder[1]?.center && centerOrder[0]?.center === b.center) return 1;
            if (centerOrder[0]?.center !== centerOrder[1]?.center && centerOrder[0]?.center === a.center) return -1;
            return String(a.id).localeCompare(String(b.id), 'ru', { numeric: true });
        });
}

function renderTargetIndicatorTaskIds(projects, mode) {
    if (!projects || projects.length === 0) {
        return `
            <details class="target-indicator-task-ids">
                <summary class="target-indicator-task-ids__title">ID задач в расчете <span>0</span></summary>
                <div class="target-indicator-task-ids__empty">Нет задач</div>
            </details>
        `;
    }

    if (mode === 'all') {
        const byCenter = new Map();
        projects.forEach(project => {
            const group = getTargetIndicatorDisplayGroup(project.center);
            if (!group) return;
            if (!byCenter.has(group.key)) {
                byCenter.set(group.key, {
                    center: group.key,
                    centerName: group.label,
                    ids: []
                });
            }
            byCenter.get(group.key).ids.push(project.id);
        });

        const rows = getOrderedTargetIndicatorRows(Array.from(byCenter.values())).map(row => `
            <div class="target-indicator-task-ids__group">
                <div class="target-indicator-task-ids__group-title">${escapeHtml(row.centerName || row.center)} <span>${escapeHtml(String(row.ids.length))}</span></div>
                <div class="target-indicator-task-ids__chips">
                    ${row.ids.map(id => `<span class="target-indicator-task-id">${escapeHtml(id)}</span>`).join('')}
                </div>
            </div>
        `).join('');

        return `
            <details class="target-indicator-task-ids">
                <summary class="target-indicator-task-ids__title">ID задач в расчете <span>${escapeHtml(String(projects.length))}</span></summary>
                ${rows}
            </details>
        `;
    }

    return `
        <details class="target-indicator-task-ids">
            <summary class="target-indicator-task-ids__title">ID задач в расчете <span>${escapeHtml(String(projects.length))}</span></summary>
            <div class="target-indicator-task-ids__chips">
                ${projects.map(project => `<span class="target-indicator-task-id">${escapeHtml(project.id)}</span>`).join('')}
            </div>
        </details>
    `;
}

function renderTargetIndicatorMetric(label, percent, value, total, options = {}) {
    const extraNote = getCleanTextValue(options.extraNote);

    return `
        <div class="target-indicator-metric">
            <span class="target-indicator-metric__label">${escapeHtml(label)}</span>
            <span class="target-indicator-metric__value">${escapeHtml(formatTargetIndicatorPercent(percent))}</span>
            <span class="target-indicator-metric__note">${escapeHtml(String(value))} из ${escapeHtml(String(total))}</span>
            ${extraNote ? `<span class="target-indicator-metric__extra">${escapeHtml(extraNote)}</span>` : ''}
        </div>
    `;
}

function formatTargetIndicatorPercent(value) {
    return value === null ? '-' : `${value}%`;
}

function updateIndicatorsLegend() {
    try {
        const legendItems = document.getElementById('legend-items');
        if (!legendItems) return;
        
        const legendData = [
            { tone: 'success', name: '25-100%', description: 'запас часов / резерв' },
            { tone: 'warning', name: '0-24%', description: 'приближение к лимиту' },
            { tone: 'danger', name: '<0%', description: 'перерасход' },
            { tone: 'neutral', name: 'нет данных', description: 'нет плановых часов' }
        ];
        
        legendItems.innerHTML = legendData.map(item => `
            <div class="legend-item">
                <span class="project-value-badge project-value-badge--${escapeHtml(item.tone)} project-value-badge--legend">${escapeHtml(item.name)}</span>
                <div class="legend-text">${escapeHtml(item.description)}</div>
            </div>
        `).join('');
        
        const legend = document.getElementById('indicators-legend');
        if (legend) legend.style.display = 'flex';
    } catch (error) {
        console.warn('Ошибка в updateIndicatorsLegend:', error);
    }
}

function updateEmployeeCounts() {
    try {
        const totalCount = parseInt(getCellValue(EXCEL_CONFIG.header.employeeCount, worksheetCache)) || 0;
        const headerCountEl = document.getElementById('header-employee-count');
        const countValueEl = document.getElementById('total-employees-count');
        
        if (headerCountEl && countValueEl) {
            if (totalCount > 0) {
                countValueEl.textContent = totalCount;
                headerCountEl.style.display = 'inline-flex';
            } else {
                headerCountEl.style.display = 'none';
            }
        }
        
        const workloadTotal = parseInt(getCellValue(EXCEL_CONFIG.header.workloadTotal, worksheetCache)) || 0;
        const workloadEl = document.getElementById('workload-count-badge');
        const workloadValueEl = document.getElementById('workload-total-count');
        
        if (workloadEl && workloadValueEl) {
            if (workloadTotal > 0) {
                workloadValueEl.textContent = workloadTotal;
                workloadEl.style.display = 'inline-flex';
            } else {
                workloadEl.style.display = 'none';
            }
        }
    } catch (error) {
        console.warn('Ошибка в updateEmployeeCounts:', error);
    }
}

function createTodayAbsenceEmployeeMaps() {
    const employeeMaps = {};

    TODAY_ABSENCE_ORDER.forEach(statusKey => {
        employeeMaps[statusKey] = new Map();
    });

    return employeeMaps;
}

function buildTodayAbsenceDataFromMaps(employeeMaps, referenceDate) {
    const categories = {};
    const totalEmployees = new Map();

    TODAY_ABSENCE_ORDER.forEach(statusKey => {
        const map = employeeMaps[statusKey] || new Map();

        map.forEach((name, employeeKey) => {
            if (!totalEmployees.has(employeeKey)) {
                totalEmployees.set(employeeKey, name);
            }
        });

        const employees = Array.from(map.values())
            .filter(Boolean)
            .sort((a, b) => a.localeCompare(b, 'ru', { numeric: true, sensitivity: 'base' }));

        categories[statusKey] = {
            label: TODAY_ABSENCE_LABELS[statusKey],
            count: employees.length,
            employees
        };
    });

    return {
        categories,
        total: totalEmployees.size,
        referenceDate
    };
}

function emptyTodayAbsenceData(referenceDate = getTodayCalendarDate()) {
    return buildTodayAbsenceDataFromMaps(createTodayAbsenceEmployeeMaps(), referenceDate);
}

function addTodayAbsenceEmployee(employeeMaps, statusKey, employeeName, rowIndex) {
    const map = employeeMaps[statusKey];
    if (!map) return;

    const displayName = getCleanTextValue(employeeName);
    const employeeKey = normalizePersonName(displayName) || `row-${rowIndex}`;

    if (!map.has(employeeKey)) {
        map.set(employeeKey, displayName || `Строка ${rowIndex + 1}`);
    }
}

const TODAY_ABSENCE_KNOWN_CENTERS = new Set(['ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ', 'УВП', 'ЛОИ']);

function getTodayAbsenceCenterFilter() {
    const normalized = normalizeProjectCenterLabel(currentSheet);
    return TODAY_ABSENCE_KNOWN_CENTERS.has(normalized) ? normalized : '';
}

function getTodayAbsenceDataFromTechBlockLoad() {
    const referenceDate = getTodayCalendarDate();
    const context = getWorksheetHeaderContext('Загрузка_ТехБлок');

    if (!context) {
        return emptyTodayAbsenceData(referenceDate);
    }

    const centerFilter = getTodayAbsenceCenterFilter();
    const employeeLookup = getTechBlockEmployeeLookup();
    const employeeMaps = createTodayAbsenceEmployeeMaps();

    for (let rowIndex = context.range.s.r + 1; rowIndex <= context.range.e.r; rowIndex++) {
        const rawId = getWorksheetCellByColumnLetter(context, rowIndex, 'I');
        if (String(rawId || '').trim() !== '-') continue;

        const rowDate = parseDateValue(getWorksheetCellByColumnLetter(context, rowIndex, 'B'));
        if (!isSameCalendarDate(rowDate, referenceDate)) continue;

        const employeeName = getWorksheetContextValue(context, rowIndex, [
            'ФИО',
            'Сотрудник',
            'ФИО сотрудника'
        ]);

        if (centerFilter) {
            const employee = employeeLookup.get(normalizePersonName(getCleanTextValue(employeeName)));
            const rowCenter = normalizeProjectCenterLabel(getWorksheetContextValue(context, rowIndex, ['Центр'])) || employee?.center || '';
            if (rowCenter !== centerFilter) continue;
        }

        let rawAbsence = getWorksheetContextValue(context, rowIndex, ['Отсутствия']);
        if (!isValid(rawAbsence)) {
            rawAbsence = getWorksheetCellByColumnLetter(context, rowIndex, 'M');
        }

        const statusKey = TODAY_ABSENCE_STATUS_MAP[normalizeAbsenceCode(rawAbsence)];
        if (!statusKey) continue;

        addTodayAbsenceEmployee(employeeMaps, statusKey, employeeName, rowIndex);
    }

    return buildTodayAbsenceDataFromMaps(employeeMaps, referenceDate);
}

function updateAbsences() {
    try {
        const absenceData = getTodayAbsenceDataFromTechBlockLoad();
        let html = '';

        TODAY_ABSENCE_ORDER.forEach(statusKey => {
            const item = absenceData.categories[statusKey] || {
                label: TODAY_ABSENCE_LABELS[statusKey],
                count: 0,
                employees: []
            };
            const hasEmployees = item.employees.length > 0;

            html += `
                <div class="data-item">
                    <div class="data-header">
                        <span class="data-label">${escapeHtml(item.label)}</span>
                        <span class="data-value">${item.count}</span>
                        <button class="toggle-btn" onclick="window.toggleEmployeesList(this)">${hasEmployees ? '▼' : '−'}</button>
                    </div>
                    ${hasEmployees ? `
                        <div class="employees-list">
                            ${item.employees.map(emp => `<div class="employee-item">${escapeHtml(emp)}</div>`).join('')}
                        </div>
                    ` : ''}
                </div>
            `;
        });
        
        html += `
            <div class="data-item total">
                <div class="data-header data-header--static">
                    <span class="data-label data-label_Total">Всего отсутствий</span>
                    <span class="data-value">${absenceData.total}</span>
                    <span class="data-toggle-placeholder" aria-hidden="true"></span>
                </div>
            </div>
        `;
        
        const container = document.getElementById('absences-container');
        if (container) container.innerHTML = html;
        
        updateAbsencesHeader(absenceData.referenceDate);
    } catch (error) {
        console.warn('Ошибка в updateAbsences:', error);
    }
}

function updateAbsencesHeader(date = getTodayCalendarDate()) {
    try {
        const absencesBlock = document.getElementById('absences-block');
        if (!absencesBlock) return;
        
        const h3 = absencesBlock.querySelector('h3');
        if (!h3) return;
        
        const headerDate = date instanceof Date && !isNaN(date.getTime()) ? date : getTodayCalendarDate();
        const dateString = headerDate.toLocaleDateString('ru-RU', {
            day: 'numeric',
            month: 'long',
            year: 'numeric'
        });
        
        h3.textContent = `Табель ${dateString}`;
    } catch (error) {
        console.warn('Ошибка в updateAbsencesHeader:', error);
    }
}

function updateLoadingAttention() {
    try {
        const config = EXCEL_CONFIG.loading;
        
        const loaded = parseInt(getCellValue(config.loadedCell, worksheetCache)) || 0;
        const notLoaded = parseInt(getCellValue(config.notLoadedCell, worksheetCache)) || 0;
        const employees = getEmployeesList(config.employeesColumn, config.employeesStartRow, worksheetCache);
        
        let html = '';
        
        if (employees.length > 0) {
            html += `
                <div class="data-item">
                    <div class="data-header">
                        <span class="data-label">Требуют внимания</span>
                        <span class="data-value">${notLoaded}</span>
                        <button class="toggle-btn" onclick="window.toggleEmployeesList(this)">▼</button>
                    </div>
                    <div class="employees-list">
                        ${employees.map(emp => `<div class="employee-item">${escapeHtml(emp)}</div>`).join('')}
                    </div>
                </div>
            `;
        } else {
            html += `
                <div class="data-item">
                    <div class="data-header data-header--static">
                        <span class="data-label">Требуют внимания</span>
                        <span class="data-value">${notLoaded}</span>
                        <span class="data-toggle-placeholder" aria-hidden="true"></span>
                    </div>
                </div>
            `;
        }
        
        html += `
            <div class="data-item total">
                <div class="data-header data-header--static">
                    <span class="data-label data-label_Total">Всего загружено человек</span>
                    <span class="data-value">${loaded}</span>
                    <span class="data-toggle-placeholder" aria-hidden="true"></span>
                </div>
            </div>
        `;
        
        const container = document.getElementById('loading-container');
        if (container) container.innerHTML = html;
    } catch (error) {
        console.warn('Ошибка в updateLoadingAttention:', error);
    }
}

function updateWorkloadAttentionBadge(notificationData) {
    const badgeEl = document.getElementById('workload-attention-badge');
    if (!badgeEl) return;

    if (notificationData && notificationData.hasNotification) {
        badgeEl.textContent = `Требует внимания • ${notificationData.attentionCount}`;
        badgeEl.style.display = 'inline-flex';
    } else {
        badgeEl.style.display = 'none';
    }
}

function renderChartCards() {
    try {
        const periods = EXCEL_CONFIG.charts.periods;
        const chartsGrid = document.getElementById('charts-grid');
        
        if (!chartsGrid) {
            console.warn('Элемент charts-grid не найден');
            return;
        }
        
        const managementNames = EXCEL_CONFIG.managements.columns.map(col => 
            getCellValue(col + EXCEL_CONFIG.managements.nameRow, worksheetCache) || ''
        );
        
        const notificationData = getNotificationData();
        const hasNotification = notificationData.hasNotification;
        updateWorkloadAttentionBadge(notificationData);
        
        // Получаем данные для расчёта разницы между месяцами
const currentMonthData = periods.find(p => p.id === 'month');
const lastMonthData = periods.find(p => p.id === 'last-month');

let currentDiffPercent = null;   // для текущего месяца (к прошлому)
let lastDiffPercent = null;       // для прошлого месяца (к позапрошлому)

// Расчёт для текущего месяца (к прошлому)
if (currentMonthData && lastMonthData) {
    const currentLoad = parseFloat(getCellValue(currentMonthData.loadCell, worksheetCache)) || 0;
    const currentTotal = parseFloat(getCellValue(currentMonthData.totalCell, worksheetCache)) || 0;
    const lastLoad = parseFloat(getCellValue(lastMonthData.loadCell, worksheetCache)) || 0;
    const lastTotal = parseFloat(getCellValue(lastMonthData.totalCell, worksheetCache)) || 0;
    
    const currentPercentage = currentTotal > 0 ? (currentLoad / currentTotal) * 100 : 0;
    const lastPercentage = lastTotal > 0 ? (lastLoad / lastTotal) * 100 : 0;
    
    if (!isNaN(currentPercentage) && !isNaN(lastPercentage) && lastPercentage > 0) {
        currentDiffPercent = ((currentPercentage - lastPercentage) / lastPercentage) * 100;
    }
}

// Расчёт для прошлого месяца (к позапрошлому) - данные из ячеек N64 и N65
if (lastMonthData) {
    const lastLoad = parseFloat(getCellValue(lastMonthData.loadCell, worksheetCache)) || 0;
    const lastTotal = parseFloat(getCellValue(lastMonthData.totalCell, worksheetCache)) || 0;
    
    // Данные позапрошлого месяца из ячеек N64 (загрузка) и N65 (всего часов)
    const beforeLastLoad = parseFloat(getCellValue('N64', worksheetCache)) || 0;
    const beforeLastTotal = parseFloat(getCellValue('N65', worksheetCache)) || 0;
    
    const lastPercentage = lastTotal > 0 ? (lastLoad / lastTotal) * 100 : 0;
    const beforeLastPercentage = beforeLastTotal > 0 ? (beforeLastLoad / beforeLastTotal) * 100 : 0;
    
    if (!isNaN(lastPercentage) && !isNaN(beforeLastPercentage) && beforeLastPercentage > 0) {
        lastDiffPercent = ((lastPercentage - beforeLastPercentage) / beforeLastPercentage) * 100;
    }
}
        
        let cardsHTML = '';
        
        periods.forEach((period, index) => {
            const periodTitle = getCellValue(period.titleCell, worksheetCache) || period.title;
            const subtitle = getCellValue(period.subtitleCell, worksheetCache);
            const { types, values } = getProjectTypesAndValues(index);
            const percentages = getManagementPercentagesForPeriod(period);
            const managementSectionTitle = currentSheet === 'ТехБлок'
                ? 'В разрезе центров:'
                : 'В разрезе управлений:';
            
            let periodClass = '';
            if (period.id === 'day' || period.id === 'last-week' || period.id === 'month') {
                periodClass = 'current-period';
            } else if (period.id === 'last-month') {
                periodClass = 'past-period';
            } else if (period.id === 'next-month') {
                periodClass = 'future-period';
            }
            
            const maxValue = values.length > 0 ? Math.max(...values) : 0;
            
            const isMonthWithAttention = period.id === 'month' && hasNotification;
            
            // Определяем, какой diff показывать для текущего периода
            // Определяем, какой diff показывать для текущего периода
	let diffHtml = '';
	if (period.id === 'month' && currentDiffPercent !== null) {
	    diffHtml = `
	        <div class="chart-diff ${currentDiffPercent >= 0 ? 'diff-positive' : 'diff-negative'}">
	            ${currentDiffPercent >= 0 ? '▲' : '▼'} ${Math.abs(currentDiffPercent).toFixed(1)}%
	        </div>
	    `;
	} else if (period.id === 'last-month' && lastDiffPercent !== null) {
	    diffHtml = `
	        <div class="chart-diff ${lastDiffPercent >= 0 ? 'diff-positive' : 'diff-negative'}">
	            ${lastDiffPercent >= 0 ? '▲' : '▼'} ${Math.abs(lastDiffPercent).toFixed(1)}%
	        </div>
	    `;
	}
            
            cardsHTML += `
                <div class="chart-card ${periodClass} ${isMonthWithAttention ? 'chart-card--attention' : ''}" id="${period.id}-card">
                    <div class="chart-card__summary">
                        <div class="chart-title-row">
                            <h4 class="chart-title">${escapeHtml(periodTitle)}</h4>
                        </div>
                        <h4 class="chart-title-date" id="${period.id}-chart-title">${escapeHtml(subtitle || '')}</h4>
                        
                        <div class="chart-content-row">
                            <div class="chart-left">
	                                <div class="chart-wrapper">
	                                    <canvas id="${period.id}-chart" width="140" height="140"></canvas>
	                                    <div id="${period.id}-percentage" class="chart-percentage">0%</div>
	                                    ${diffHtml}
	                                </div>
                            </div>
                            <div class="chart-right">
                                <div class="chart-hours-heading-row" id="${period.id}-hours-heading-row">
                                    <div class="chart-hours-heading">Загрузка в часах</div>
                                    <div id="${period.id}-overtime" class="chart-hours-overtime-badge" style="display: none;"></div>
                                </div>
                                <div id="${period.id}-hours" class="chart-hours-side"></div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="project-types-info">
                        <div class="management-name-title">В разрезе видов проектов:</div>
                        ${types.map((type, i) => `
                            <div class="project-type-item">
                                <span>${escapeHtml(type)}</span>
                                <span ${values[i] === maxValue && maxValue > 0 ? 'style="font-weight: bold;"' : ''}>${values[i]}%</span>
                            </div>
                        `).join('')}
                    </div>
                    
                    <div class="management-info">
                        <div class="management-name-title">${managementSectionTitle}</div>
                        ${managementNames.map((name, i) => {
                            if (!name) return '';
                            const percentageValue = percentages[i] || 0;
                            const color = getManagementGaugeColor(percentageValue);
                            return `
                                <div class="management-item">
                                    <span class="management-name">${escapeHtml(name)}</span>
                                    <span class="management-percentage" style="background-color: ${color}20; color: ${color};">${Math.round(percentageValue)}%</span>
                                </div>
                            `;
                        }).join('')}
                    </div>
                </div>
            `;
        });
        
        chartsGrid.innerHTML = cardsHTML;

        // Даём браузеру отрендерить canvas перед созданием графиков
        requestAnimationFrame(() => updateChartsData());

    } catch (error) {
        console.error('Ошибка в renderChartCards:', error);
    }
}

function renderCenterMonthlyLoadSection() {
    const container = document.getElementById('center-monthly-load-section');
    if (!container) return;

    const centerKey = normalizeProjectCenterLabel(currentSheet) || currentSheet;
    const normalizedCenter = normalizeProjectCenterLabel(currentSheet);

    if (normalizedCenter) {
        const quarters = getQuarterlyTrendDisplayQuarters(
            getCenterQuarterlyTrendDataFromTables(normalizedCenter),
            { holdLastFutureDrop: true }
        );

        if (quarters.length === 0) {
            chartManager.destroy('center-quarterly-trend-chart');
            container.innerHTML = '';
            container.style.display = 'none';
            return;
        }

        container.style.display = 'block';
        container.innerHTML = renderQuarterlyTrendMarkup(quarters, 'center-quarterly-trend-chart');

        requestAnimationFrame(() => createOrUpdateQuarterlyTrendChart(quarters, 'center-quarterly-trend-chart'));
        return;
    }

    chartManager.destroy('center-quarterly-trend-chart');

    if (currentSheet === 'ТехБлок' || currentSheet === 'НКиТК') {
        container.innerHTML = '';
        container.style.display = 'none';
        return;
    }

    const baseMonths = getCenterMonthlyLoadMonths();
    const projectHoursByMonth = getSpentProjectHoursByMonthForCenter(centerKey);
    const forecastHoursByMonth = getForecastProjectHoursByMonthForCenter(centerKey);
    const months = getCenterMonthlyLoadMonthsWithData(centerKey, baseMonths, projectHoursByMonth, forecastHoursByMonth);
    if (months.length === 0) {
        container.innerHTML = '';
        container.style.display = 'none';
        return;
    }

    const loadMonths = getMonthlyProjectLoadRows(months, projectHoursByMonth, forecastHoursByMonth, {
        showAllForecastHours: true
    });
    const currentMonth = getCurrentCenterMonthlyLoadRecord(loadMonths);
    const nextForecastMonth = getNextForecastMonthlyLoadRecord(loadMonths);
    const currentMonthKey = `${currentMonth.year}-${String(currentMonth.monthIndex + 1).padStart(2, '0')}`;
    const currentProjectHours = projectHoursByMonth.get(currentMonthKey) || 0;
    const currentProjectPercent = currentMonth.workPercent;
    const nextForecastPercent = nextForecastMonth ? nextForecastMonth.forecastPercent : null;
    const latestPercent = currentProjectPercent === null ? '-' : `${currentProjectPercent}%`;
    container.style.display = 'block';
    container.innerHTML = `
        <div class="center-monthly-load-card" data-fact-toggle-block="center-monthly-load">
            <div class="center-monthly-load-card__header">
                <div>
                    <h4 class="center-monthly-load-card__title">Общая загрузка по месяцам</h4>
                    <p class="center-monthly-load-card__subtitle"><span data-fact-field>Показывает нагрузку центра по часам во времени: где центр уже был загружен фактическими списаниями, </span>и где впереди ожидается нагрузка относительно доступного фонда часов.</p>
                </div>
                <div class="center-monthly-load-card__actions">
                    <div class="center-monthly-load-card__badge" data-fact-field>${escapeHtml(latestPercent)} • ${escapeHtml(currentMonth.label)}</div>
                    ${renderFactToggle('center-monthly-load-fact', true)}
                </div>
            </div>
            <div class="center-monthly-load-kpis">
                ${renderCenterMonthlyLoadKpi('Факт часов', formatHoursValue(currentProjectHours), 'ч', 'green', 'data-fact-field')}
                ${renderCenterMonthlyLoadKpi('Фонд часов', formatHoursValue(currentMonth.totalHours), 'ч', 'blue')}
                ${renderCenterMonthlyLoadKpi('Факт', currentProjectPercent === null ? '-' : `${currentProjectPercent}%`, '', 'green', 'data-fact-field')}
                ${renderCenterMonthlyLoadKpi('Прогноз / очередь', nextForecastPercent === null ? '-' : `${nextForecastPercent}%`, nextForecastMonth ? nextForecastMonth.label : '', 'blue')}
            </div>
            <div class="center-monthly-load-heatwrap">
                <div class="center-monthly-load-heat-title">
                    Общая загрузка
                    <small><span data-fact-field>факт и </span>будущая потребность</small>
                </div>
                ${renderMonthlyLoadScale(loadMonths, { workLabel: 'факт' })}
            </div>
        </div>
    `;
}

function getCurrentCenterMonthlyLoadRecord(months) {
    const referenceDate = getReferenceDate();
    const current = months.find(month => (
        month.year === referenceDate.getFullYear()
        && month.monthIndex === referenceDate.getMonth()
    ));

    if (current) return current;

    for (let index = months.length - 1; index >= 0; index--) {
        const month = months[index];
        if (month.totalHours > 0 || month.loadHours > 0) return month;
    }

    return months[months.length - 1];
}

function renderFactToggle(id, checked = false) {
    return `
        <label class="fact-toggle" title="Показать или скрыть фактические данные">
            <input class="fact-toggle__input" id="${escapeHtml(id)}" type="checkbox"${checked ? ' checked' : ''}>
            <span class="fact-toggle__box" aria-hidden="true"></span>
            <span class="fact-toggle__label">Факт</span>
        </label>
    `;
}

function handleFactToggleChange(input) {
    const block = input.closest('[data-fact-toggle-block]');
    if (!block) return;

    block.classList.toggle('fact-hidden', !input.checked);
}

function renderCenterMonthlyLoadKpi(title, value, note, tone, extraAttributes = '') {
    return `
        <div class="center-monthly-load-kpi center-monthly-load-kpi--${tone}" ${extraAttributes}>
            <span class="center-monthly-load-kpi__title">${escapeHtml(title)}</span>
            <span class="center-monthly-load-kpi__value">${escapeHtml(String(value))}</span>
            ${note ? `<span class="center-monthly-load-kpi__note">${escapeHtml(String(note))}</span>` : ''}
        </div>
    `;
}

function formatPercentPointDiff(value) {
    if (!Number.isFinite(value)) return '-';
    const rounded = Math.round(value);
    if (rounded === 0) return '0 п.п.';
    return `${rounded > 0 ? '+' : ''}${rounded} п.п.`;
}

function getActualProjectHoursByMonthForCenter(centerKey, projects = null) {
    const result = new Map();
    if (!centerKey || !workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) return result;

    const sourceProjects = projects || getTechBlockProjectPlanFactDataFromTables();

    sourceProjects
        .forEach(project => {
            const centerMonthMap = project.actualHoursByCenterMonth instanceof Map
                ? project.actualHoursByCenterMonth.get(centerKey)
                : null;

            if (!(centerMonthMap instanceof Map)) return;
            centerMonthMap.forEach((value, monthKey) => addMonthMapValue(result, monthKey, value));
        });

    return result;
}

function getSpentProjectHoursByMonthForCenter(centerKey) {
    const result = new Map();
    if (!centerKey || !workbook || !workbook.Sheets || !workbook.Sheets['Загрузка_ТехБлок']) return result;

    const context = getWorksheetHeaderContext('Загрузка_ТехБлок');
    const employeeLookup = getTechBlockEmployeeLookup();

    if (!context) return result;

    for (let rowIndex = context.range.s.r + 1; rowIndex <= context.range.e.r; rowIndex++) {
        const date = parseDateValue(getWorksheetContextValue(context, rowIndex, ['Дата']));
        if (!(date instanceof Date) || isNaN(date.getTime())) continue;

        const fio = getCleanTextValue(getWorksheetContextValue(context, rowIndex, ['ФИО']));
        const employee = employeeLookup.get(normalizePersonName(fio));
        const center = normalizeProjectCenterLabel(getWorksheetContextValue(context, rowIndex, ['Центр'])) || employee?.center || '';
        if (center !== centerKey) continue;

        const hours = parseNumericValue(getWorksheetContextValue(context, rowIndex, ['Часы'])) || 0;
        if (!hours) continue;

        const absence = getCleanTextValue(getWorksheetContextValue(context, rowIndex, ['Отсутствия']));
        const isAbsence = isProjectAbsenceValue(absence);
        const isForecast = isForecastLoadValue(getWorksheetContextValue(context, rowIndex, ['Прогноз']));

        if (isAbsence || isForecast) continue;

        addMonthMapValue(result, getMonthKey(date), hours);
    }

    return result;
}

function getActiveProjectActualHoursByMonth(projects = null) {
    const result = new Map();
    if (!workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) return result;

    const sourceProjects = projects || getTechBlockProjectPlanFactDataFromTables();

    sourceProjects
        .filter(project => project.status && project.status.type === 'active')
        .forEach(project => {
            (project.actualHoursByCenterMonth || new Map()).forEach(monthMap => {
                if (!(monthMap instanceof Map)) return;
                monthMap.forEach((value, monthKey) => addMonthMapValue(result, monthKey, value));
            });
        });

    return result;
}

function getForecastProjectHoursByMonthForCenter(centerKey) {
    const totalHoursByMonth = new Map();
    const salesPlanHoursByMonth = new Map();
    const loadForecastHoursByMonth = new Map();
    const result = {
        totalHoursByMonth,
        salesPlanHoursByMonth,
        loadForecastHoursByMonth
    };
    if (!centerKey || !workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) return result;

    const projectRows = getWorksheetRowsByHeaders('Проекты_ТехБлок');
    const managementCenterLookup = getTechBlockManagementCenterLookup(projectRows);
    const salesPlanByProject = getTechBlockSalesPlanByProject(managementCenterLookup);

    salesPlanByProject.forEach(project => {
        const planMonthMap = project.planHoursByCenterMonth instanceof Map
            ? project.planHoursByCenterMonth.get(centerKey)
            : null;

        if (!(planMonthMap instanceof Map)) return;

        planMonthMap.forEach((value, monthKey) => {
            addMonthMapValue(salesPlanHoursByMonth, monthKey, value);
            addMonthMapValue(totalHoursByMonth, monthKey, value);
        });
    });

    return result;
}

function getForecastProjectHoursByMonth(projects = null) {
    const result = new Map();
    if (!workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) return result;

    const sourceProjects = projects || getTechBlockProjectPlanFactDataFromTables();

    sourceProjects
        .filter(project => project.status && project.status.type !== 'done' && project.status.type !== 'cancelled')
        .forEach(project => {
            (project.planHoursByCenterMonth || new Map()).forEach(monthMap => {
                if (monthMap instanceof Map) {
                    monthMap.forEach((value, monthKey) => addMonthMapValue(result, monthKey, value));
                }
            });
        });

    return result;
}

function hasMonthMapHours(monthMap) {
    if (!(monthMap instanceof Map)) return false;
    for (const value of monthMap.values()) {
        if ((value || 0) > 0) return true;
    }
    return false;
}

function hasNestedMonthMapHours(nestedMap) {
    if (!(nestedMap instanceof Map)) return false;
    for (const monthMap of nestedMap.values()) {
        if (hasMonthMapHours(monthMap)) return true;
    }
    return false;
}

function getCenterMonthlyLoadMonths() {
    const currentYear = getReferenceDate().getFullYear();

    return getQuarterlyTrendMonths()
        .filter(record => (
            record.year >= currentYear
            && Number.isFinite(record.loadHours)
            && Number.isFinite(record.totalHours)
            && (record.loadHours > 0 || record.totalHours > 0 || Number.isFinite(record.percentage))
        ))
        .map(record => {
            const calculatedPercent = record.totalHours > 0 ? (record.loadHours / record.totalHours) * 100 : null;
            const percent = Number.isFinite(calculatedPercent)
                ? calculatedPercent
                : Number.isFinite(record.percentage)
                    ? record.percentage
                    : null;
            const loadPercent = percent === null ? null : Math.round(percent);

            return {
                ...record,
                label: record.label || formatMonthName(record.monthIndex),
                loadPercent
            };
        });
}

function getCenterMonthlyLoadMonthsWithData(centerKey, baseMonths, ...monthSources) {
    const currentYear = getReferenceDate().getFullYear();
    const monthByKey = new Map();

    (baseMonths || []).forEach(month => {
        const monthKey = `${month.year}-${String(month.monthIndex + 1).padStart(2, '0')}`;
        monthByKey.set(monthKey, month);
    });

    const dataMonthKeys = new Set();
    monthSources.forEach(source => collectMonthlyLoadSourceKeys(source, dataMonthKeys));

    const missingMonthKeys = Array.from(dataMonthKeys)
        .filter(monthKey => !monthByKey.has(monthKey))
        .filter(monthKey => {
            const date = parseMonthKeyToDate(monthKey);
            return date && date.getFullYear() >= currentYear;
        })
        .sort();

    if (missingMonthKeys.length === 0) {
        return Array.from(monthByKey.values()).sort(sortMonthlyLoadRecords);
    }

    const monthIntervals = missingMonthKeys
        .map(monthKey => parseMonthKeyToDate(monthKey))
        .filter(Boolean)
        .map(monthStart => ({
            start: monthStart,
            end: new Date(monthStart.getFullYear(), monthStart.getMonth() + 1, 0)
        }));
    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);
    const centerCapacityByMonth = capacityByCenterMonth.get(centerKey) || new Map();

    missingMonthKeys.forEach(monthKey => {
        const monthStart = parseMonthKeyToDate(monthKey);
        if (!monthStart) return;

        monthByKey.set(monthKey, {
            year: monthStart.getFullYear(),
            monthIndex: monthStart.getMonth(),
            label: formatMonthName(monthStart.getMonth()),
            loadHours: 0,
            totalHours: centerCapacityByMonth.get(monthKey)?.availableHours || 0,
            percentage: null,
            loadPercent: null
        });
    });

    return Array.from(monthByKey.values()).sort(sortMonthlyLoadRecords);
}

function collectMonthlyLoadSourceKeys(source, target) {
    if (!source || !target) return;

    if (source instanceof Map) {
        source.forEach((value, monthKey) => {
            if ((value || 0) > 0) target.add(monthKey);
        });
        return;
    }

    ['totalHoursByMonth', 'salesPlanHoursByMonth', 'loadForecastHoursByMonth'].forEach(field => {
        if (!(source[field] instanceof Map)) return;
        source[field].forEach((value, monthKey) => {
            if ((value || 0) > 0) target.add(monthKey);
        });
    });
}

function sortMonthlyLoadRecords(a, b) {
    if (a.year !== b.year) return a.year - b.year;
    return a.monthIndex - b.monthIndex;
}

function getCenterMonthlyLoadMonthsForRange(startDate, endDate) {
    const monthIntervals = getPlanFactMonthIntervals(startDate, endDate);
    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);
    const capacityByMonth = new Map();

    capacityByCenterMonth.forEach(monthMap => {
        monthMap.forEach((value, monthKey) => {
            addMonthMapValue(capacityByMonth, monthKey, value.availableHours || 0);
        });
    });

    return monthIntervals.map(interval => {
        const monthKey = getMonthKey(interval.start);
        const totalHours = capacityByMonth.get(monthKey) || 0;

        return {
            year: interval.start.getFullYear(),
            monthIndex: interval.start.getMonth(),
            label: formatMonthName(interval.start.getMonth()),
            loadHours: 0,
            totalHours,
            percentage: null,
            loadPercent: null
        };
    });
}

function getMonthlyProjectLoadRows(months, workHoursByMonth, forecastHoursByMonth, options = {}) {
    const nextMonthStart = getNextMonthStartDate();
    const showAllForecastHours = options.showAllForecastHours === true;
    const forecastTotalHoursByMonth = forecastHoursByMonth instanceof Map
        ? forecastHoursByMonth
        : forecastHoursByMonth?.totalHoursByMonth instanceof Map
            ? forecastHoursByMonth.totalHoursByMonth
            : new Map();
    const salesPlanHoursByMonth = forecastHoursByMonth?.salesPlanHoursByMonth instanceof Map
        ? forecastHoursByMonth.salesPlanHoursByMonth
        : new Map();
    const loadForecastHoursByMonth = forecastHoursByMonth?.loadForecastHoursByMonth instanceof Map
        ? forecastHoursByMonth.loadForecastHoursByMonth
        : new Map();

    return months.map(month => {
        const monthKey = `${month.year}-${String(month.monthIndex + 1).padStart(2, '0')}`;
        const monthStart = new Date(month.year, month.monthIndex, 1);
        const totalHours = month.totalHours || 0;
        const workHours = workHoursByMonth instanceof Map ? (workHoursByMonth.get(monthKey) || 0) : 0;
        const rawForecastHours = forecastTotalHoursByMonth.get(monthKey) || 0;
        const forecastHours = showAllForecastHours || monthStart >= nextMonthStart ? rawForecastHours : 0;
        const forecastSalesPlanHours = showAllForecastHours || monthStart >= nextMonthStart
            ? (salesPlanHoursByMonth.get(monthKey) || 0)
            : 0;
        const forecastLoadHours = showAllForecastHours || monthStart >= nextMonthStart
            ? (loadForecastHoursByMonth.get(monthKey) || 0)
            : 0;

        return {
            ...month,
            monthKey,
            workHours,
            forecastHours,
            forecastSalesPlanHours,
            forecastLoadHours,
            workPercent: totalHours > 0 ? Math.round((workHours / totalHours) * 100) : null,
            forecastPercent: totalHours > 0 ? Math.round((forecastHours / totalHours) * 100) : null
        };
    });
}

function getNextMonthStartDate() {
    const referenceDate = getReferenceDate();
    return new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 1);
}

function getNextForecastMonthlyLoadRecord(months) {
    const nextMonthStart = getNextMonthStartDate();

    return months.find(month => {
        const monthStart = new Date(month.year, month.monthIndex, 1);
        return monthStart >= nextMonthStart && (month.forecastHours || 0) > 0;
    }) || months.find(month => {
        const monthStart = new Date(month.year, month.monthIndex, 1);
        return monthStart >= nextMonthStart && month.forecastPercent !== null;
    }) || null;
}

function renderMonthlyLoadScale(months, options = {}) {
    const monthCount = Math.max(1, months.length);
    const forecastLabel = options.forecastLabel || 'прогноз / очередь';
    const workLabel = options.workLabel || 'в работе';

    return `
        <div class="center-monthly-load-scale" style="--monthly-load-count: ${monthCount};">
            <div class="center-monthly-load-layer center-monthly-load-layer--work" data-fact-field>
                <div class="center-monthly-load-layer__label">${escapeHtml(workLabel)}</div>
                <div class="center-monthly-load-cells center-monthly-load-cells--work">
                    ${months.map(month => renderCenterMonthlyLoadCell(month, 'work', false, workLabel)).join('')}
                </div>
            </div>
            <div class="center-monthly-load-layer center-monthly-load-layer--forecast">
                <div class="center-monthly-load-layer__label">${escapeHtml(forecastLabel)}</div>
                <div class="center-monthly-load-cells center-monthly-load-cells--forecast">
                    ${months.map(month => renderCenterMonthlyLoadCell(month, 'forecast', true, forecastLabel)).join('')}
                </div>
            </div>
        </div>
    `;
}

function renderCenterMonthlyLoadCell(month, type = 'work', showLabel = true, typeLabel = '') {
    const isForecast = type === 'forecast';
    const label = isForecast ? (typeLabel || 'прогноз / очередь') : (typeLabel || 'в работе');
    const hours = isForecast ? (month.forecastHours || 0) : (month.workHours || 0);
    const percent = isForecast ? month.forecastPercent : month.workPercent;
    const hasHours = hours > 0 || (month.totalHours || 0) > 0;
    const percentText = percent === null ? '-' : `${percent}%`;
    const titleParts = [
        `${month.label} ${month.year}`,
        `${label} ${formatHoursValue(hours)} ч`,
        `фонд ${formatHoursValue(month.totalHours || 0)} ч`,
        `загрузка ${percentText}`
    ];
    if (isForecast) {
        titleParts.push(`из ПП ${formatHoursValue(month.forecastSalesPlanHours || 0)} ч`);
    }
    const title = titleParts.join(' / ');
    const background = isForecast
        ? getForecastLoadColor(percent, hours > 0)
        : getLoadHeatColor(percent, hasHours);

    return `
        <div class="center-monthly-load-cell" title="${escapeHtml(title)}">
            <div class="center-monthly-load-cell__bar center-monthly-load-cell__bar--${escapeHtml(type)}" style="background: ${background};">
                <span>${escapeHtml(percentText)}</span>
            </div>
            ${showLabel ? `<div class="center-monthly-load-cell__label">${escapeHtml(formatCenterMonthlyLoadLabel(month))}</div>` : ''}
        </div>
    `;
}

function formatCenterMonthlyLoadLabel(month) {
    const shortMonth = formatMonthName(month.monthIndex).slice(0, 3);
    const shortYear = String(month.year).slice(-2);
    return `${shortMonth} ${shortYear}`;
}
function renderQuarterlyTrendSection() {
    const container = document.getElementById('quarterly-trend-section');
    if (!container) return;

    if (currentSheet !== 'ТехБлок') {
        chartManager.destroy('quarterly-trend-chart');
        container.innerHTML = '';
        container.style.display = 'none';
        return;
    }

    const quarters = getQuarterlyTrendDisplayQuarters(
        getTechBlockQuarterlyTrendDataFromTables(),
        { holdLastFutureDrop: true }
    );

    if (quarters.length === 0) {
        chartManager.destroy('quarterly-trend-chart');
        container.innerHTML = '';
        container.style.display = 'none';
        return;
    }

    container.style.display = 'block';
    container.innerHTML = renderQuarterlyTrendMarkup(quarters);
    requestAnimationFrame(() => createOrUpdateQuarterlyTrendChart(quarters));
}

function renderQuarterlyTrendMarkup(quarters, chartId = 'quarterly-trend-chart') {

    if (quarters.length === 0) {
        chartManager.destroy(chartId);
        return `
            <div class="quarterly-trend-card">
                <div class="quarterly-trend-empty">
                    <strong>Квартальная динамика пока недоступна</strong>
                    <span>Добавьте месячную историю на лист центра, и график появится автоматически.</span>
                </div>
            </div>
        `;
    }

    const latestQuarter = quarters[quarters.length - 1];
    const latestPercent = Math.round(latestQuarter.percentage);

    return `
        <div class="quarterly-trend-card">
            <div class="quarterly-trend-card__header">
                <div>
                    <h4 class="quarterly-trend-card__title">Динамика загрузки по кварталам</h4>
                    <p class="quarterly-trend-card__subtitle">Процент считается по плоским таблицам: фактические часы из «Загрузка_ТехБлок» к фонду часов из сотрудников и табеля.</p>
                </div>
                <div class="quarterly-trend-card__badge">${latestPercent}% • ${escapeHtml(flattenChartLabel(latestQuarter.label))}</div>
            </div>
            <div class="quarterly-trend-card__canvas-wrap">
                <canvas id="${escapeHtml(chartId)}" class="quarterly-trend-chart"></canvas>
            </div>
        </div>
    `;
}

function updateChartsData() {
    try {
        const periods = EXCEL_CONFIG.charts.periods;
        
        periods.forEach(period => {
            const loadHours = parseFloat(getCellValue(period.loadCell, worksheetCache)) || 0;
            const totalHours = parseFloat(getCellValue(period.totalCell, worksheetCache)) || 0;
            const idleHours = parseFloat(getCellValue(period.idleCell, worksheetCache)) || 0;
            const absenceHours = parseFloat(getCellValue(period.absenceCell, worksheetCache)) || 0;
            const overtimeHours = parseFloat(getCellValue(period.overtimeCell, worksheetCache)) || 0;
            
            let percentage = parseFloat(getCellValue(period.percentageCell, worksheetCache));
            if (isNaN(percentage) || percentage === '') {
                percentage = totalHours > 0 ? (loadHours / totalHours * 100) : 0;
            }

            createOrUpdateChart(period.id, percentage, loadHours, totalHours, idleHours, absenceHours, overtimeHours);
        });
        
        log('Графики обновлены');
    } catch (error) {
        console.error('Ошибка в updateChartsData:', error);
    }
}

function createOrUpdateQuarterlyTrendChart(quarters, chartId = 'quarterly-trend-chart') {
    const canvas = document.getElementById(chartId);
    if (!canvas) {
        chartManager.destroy(chartId);
        return;
    }

    const ctx = canvas.getContext('2d');
    const lineColor = getCssVar('--accent-green-strong', '#69ab81');
    const pointColor = getCssVar('--accent-green', '#84c39a');
    const gridColor = hexToRgba(getCssVar('--accent-neutral-border', '#d9e2ec'), 0.65);
    const tickColor = getCssVar('--accent-blue-text', '#4b7faa');
    const fillColor = ctx.createLinearGradient(0, 0, 0, canvas.height || 250);
    fillColor.addColorStop(0, hexToRgba(pointColor, 0.24));
    fillColor.addColorStop(1, hexToRgba(pointColor, 0.04));

    const maxValue = quarters.reduce((result, item) => Math.max(result, item.percentage), 0);
    const yAxisMax = Math.max(100, Math.ceil(maxValue / 10) * 10);

    chartManager.destroy(chartId);

    const chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: quarters.map(item => item.label),
            datasets: [{
                data: quarters.map(item => Number(item.percentage.toFixed(1))),
                borderColor: lineColor,
                backgroundColor: fillColor,
                fill: true,
                tension: 0.28,
                borderWidth: 3,
                pointRadius: 4,
                pointHoverRadius: 6,
                pointBackgroundColor: pointColor,
                pointBorderColor: '#ffffff',
                pointBorderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: {
                mode: 'index',
                intersect: false
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    enabled: true,
                    displayColors: false,
                    callbacks: {
                        title(items) {
                            return items[0]?.label || '';
                        },
                        label(context) {
                            return `Загрузка: ${Number(context.raw).toFixed(1)}%`;
                        },
                        afterLabel(context) {
                            const quarter = quarters[context.dataIndex];
                            const loadText = Math.round(quarter.loadHours).toLocaleString('ru-RU');
                            const totalText = Math.round(quarter.totalHours).toLocaleString('ru-RU');
                            return `${loadText} / ${totalText} ч.`;
                        },
                        footer(items) {
                            const quarter = quarters[items[0]?.dataIndex];
                            return quarter ? `Месяцы: ${quarter.months.join(', ')}` : '';
                        }
                    }
                }
            },
            scales: {
                x: {
                    type: 'category',
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: tickColor,
                        font: {
                            size: 11,
                            weight: '600'
                        }
                    }
                },
                y: {
                    min: 0,
                    max: yAxisMax,
                    beginAtZero: true,
                    ticks: {
                        stepSize: 20,
                        color: '#7f8c8d',
                        callback(value) {
                            return `${value}%`;
                        }
                    },
                    grid: {
                        color: gridColor,
                        drawBorder: false
                    }
                }
            }
        }
    });

    chartManager.set(chartId, chart);
}

function createOrUpdateChart(periodId, percentage, loaded, total, idleHours, absenceHours, overtimeHours) {
    const chartId = `${periodId}-chart`;
    const canvas = document.getElementById(chartId);

    if (!canvas) {
        console.warn(`Canvas для ${chartId} не найден`);
        return;
    }

    const originalPercentage = Number(percentage) || 0;
    const clampedPercentage = Math.max(0, Math.min(100, originalPercentage));

    const ctx = canvas.getContext('2d');
    const gradient = createGradient(ctx, canvas, clampedPercentage);

    chartManager.destroy(chartId);

    try {
        const chart = new Chart(ctx, {
            type: 'doughnut',
            data: {
                datasets: [{
                    data: [clampedPercentage, 100 - clampedPercentage],
                    backgroundColor: [gradient, '#ecf0f1'],
                    borderWidth: 0,
                    cutout: '65%'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    tooltip: { enabled: false }
                },
                animation: {
                    duration: 800,
                    easing: 'easeOutQuart'
                }
            }
        });
        
        chartManager.set(chartId, chart);
        
        // ДОБАВЛЕНО: проверяем наличие уведомления
        const percentEl = document.getElementById(`${periodId}-percentage`);
        if (percentEl) {
            percentEl.textContent = Math.round(originalPercentage) + '%';
            percentEl.style.color = '';
            percentEl.style.fontWeight = '';
        }
        
        const hoursEl = document.getElementById(`${periodId}-hours`);
        if (hoursEl) {
            const factLabelByPeriod = {
                month: 'факт+прогноз',
                'next-month': 'прогноз'
            };
            const factLabel = factLabelByPeriod[periodId] || 'факт';
            const overtimeEl = document.getElementById(`${periodId}-overtime`);
            const hoursHeadingRow = document.getElementById(`${periodId}-hours-heading-row`);
            if (overtimeEl) {
                if (overtimeHours > 0) {
                    overtimeEl.textContent = `в т.ч. сверхурочно ${Math.round(overtimeHours).toLocaleString('ru-RU')} ч.`;
                    overtimeEl.style.display = '';
                    if (hoursHeadingRow) hoursHeadingRow.classList.add('chart-hours-heading-row--with-overtime');
                } else {
                    overtimeEl.textContent = '';
                    overtimeEl.style.display = 'none';
                    if (hoursHeadingRow) hoursHeadingRow.classList.remove('chart-hours-heading-row--with-overtime');
                }
            }

            const metrics = [
                { value: loaded, label: factLabel, type: 'fact' },
                { value: total, label: 'план', type: 'plan' },
                { value: idleHours, label: 'простои', type: 'idle' },
                { value: absenceHours, label: 'отсутствия', type: 'absence' }
            ];

            hoursEl.innerHTML = `
                <div class="chart-hours-metrics">
                    ${metrics.map((metric, index) => `
                        ${index > 0 ? '<span class="chart-hours-metrics__divider">|</span>' : ''}
                        <span class="chart-hours-metrics__item chart-hours-metrics__item--${metric.type}">
                            <span class="chart-hours-metrics__value">${Math.round(metric.value).toLocaleString('ru-RU')}</span>
                            <span class="chart-hours-metrics__label">${escapeHtml(metric.label)}</span>
                        </span>
                    `).join('')}
                </div>
            `;
        }
        
    } catch (error) {
        console.error(`Ошибка при создании графика ${chartId}:`, error);
    }
}

function createGradient(ctx, canvas, percentage) {
    const t = Math.min(Math.max(percentage, 0), 100) / 100;
    const gradient = ctx.createRadialGradient(
        canvas.width / 2, canvas.height / 2, 0,
        canvas.width / 2, canvas.height / 2, canvas.width / 2
    );

    function mixColor(colorA, colorB, factor) {
        return colorA.map((value, index) => (
            Math.round(value + (colorB[index] - value) * factor)
        ));
    }

    function shiftColor(color, amount) {
        return color.map(value => Math.max(0, Math.min(255, Math.round(value + amount))));
    }

    function readChartRgbVar(name, fallback) {
        const raw = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
        if (!raw) return fallback;

        const parts = raw.split(',').map(part => Number(part.trim()));
        if (parts.length !== 3 || parts.some(value => Number.isNaN(value))) {
            return fallback;
        }

        return parts.map(value => Math.max(0, Math.min(255, Math.round(value))));
    }

    const lowColor = readChartRgbVar('--chart-gradient-low', [220, 107, 88]);
    const midColor = readChartRgbVar('--chart-gradient-mid', [227, 177, 96]);
    const highColor = readChartRgbVar('--chart-gradient-high', [126, 186, 143]);

    const baseColor = t < 0.5
        ? mixColor(lowColor, midColor, t / 0.5)
        : mixColor(midColor, highColor, (t - 0.5) / 0.5);

    const innerColor = shiftColor(baseColor, 16);
    const middleColor = baseColor;
    const outerColor = shiftColor(baseColor, -14);

    gradient.addColorStop(0, `rgb(${innerColor.join(', ')})`);
    gradient.addColorStop(0.55, `rgb(${middleColor.join(', ')})`);
    gradient.addColorStop(1, `rgb(${outerColor.join(', ')})`);
    
    return gradient;
}

// ИСПРАВЛЕНИЕ: добавлена проверка на существование periodConfig
function getProjectTypesAndValues(periodIndex) {
    const typesConfig = EXCEL_CONFIG.projects.typesRange;
    const periodConfig = typesConfig.periods[periodIndex];
    
    const types = [];
    const values = [];
    
    // Получаем названия типов проектов из строки заголовков (O4:S4)
    const titleColumns = getColumnRange(typesConfig.titleStartCol, typesConfig.titleEndCol);
    titleColumns.forEach(col => {
        const typeValue = getCellValue(col + typesConfig.titleRow, worksheetCache);
        if (isValid(typeValue)) {
            types.push(String(typeValue).trim());
        } else {
            types.push('');
        }
    });
    
    // Получаем значения для указанного периода из соответствующей строки
    // ИСПРАВЛЕНИЕ: добавлена проверка на существование periodConfig
    if (periodConfig) {
        const valueColumns = getColumnRange(periodConfig.startCol, periodConfig.endCol);
        valueColumns.forEach((col, index) => {
            const value = getCellValue(col + periodConfig.valuesRow, worksheetCache);
            const numValue = parseFloat(value);
            
            if (index < types.length && types[index] !== '') {
                values.push(isNaN(numValue) ? 0 : Math.round(numValue));
            }
        });
    }
    
    // Фильтруем пустые типы
    const filteredTypes = [];
    const filteredValues = [];
    types.forEach((type, i) => {
        if (type !== '') {
            filteredTypes.push(type);
            filteredValues.push(values[i] || 0);
        }
    });
    
    return { types: filteredTypes, values: filteredValues };
}

function getManagementPercentagesForPeriod(period) {
    const percentages = [];
    
    EXCEL_CONFIG.managements.columns.forEach(col => {
        const value = parseFloat(getCellValue(col + period.percentagesRow, worksheetCache)) || 0;
        percentages.push(value);
    });

    const hasVisiblePercentage = percentages.some(value => value > 0);
    if (hasVisiblePercentage) {
        return percentages;
    }

    return EXCEL_CONFIG.managements.columns.map(col => {
        const loadRow = parseInt(period.loadCell.replace(/\D/g, ''), 10);
        const totalRow = parseInt(period.totalCell.replace(/\D/g, ''), 10);
        const load = parseFloat(getCellValue(col + loadRow, worksheetCache)) || 0;
        const total = parseFloat(getCellValue(col + totalRow, worksheetCache)) || 0;

        return total > 0 ? (load / total) * 100 : 0;
    });
}

function getManagementGaugeColor(percentage) {
    if (percentage >= 90) return INDICATOR_COLORS.GREEN;
    if (percentage >= 60) return INDICATOR_COLORS.YELLOW;
    return INDICATOR_COLORS.RED;
}

// ============================================================
// ИСПРАВЛЕННАЯ ФУНКЦИЯ updateProjects С РАСКРЫТИЕМ СПИСКОВ
// ============================================================

function updateProjects() {
    try {
        const projectsSection = document.querySelector('.block-projects');
        const indicatorsLegend = document.getElementById('indicators-legend');
        const targetFilterSlot = document.getElementById('projects-target-filter-slot');
        const projectsContainer = document.getElementById('projects-container');

        if (projectsSection) projectsSection.style.display = '';

        const projectsTitle = document.querySelector('#projects-block > h3');
        if (projectsTitle) projectsTitle.textContent = 'Проекты';

        if (indicatorsLegend) indicatorsLegend.style.display = '';

        const targetIndicatorFilterAvailable = isProjectsTargetIndicatorFilterAvailable();
        if (!targetIndicatorFilterAvailable && projectsTargetIndicatorOnly) {
            projectsTargetIndicatorOnly = false;
        }

        const targetProjectKeys = targetIndicatorFilterAvailable && projectsTargetIndicatorOnly
            ? getTargetIndicatorProjectKeySet(targetIndicatorPeriodKey, { completedOnly: true })
            : null;
        const projectDetails = getProjectsFromDetails({ targetProjectKeys });
        const projectsByStatusAndType = projectDetails.projectsByKey;
        const statuses = projectDetails.statuses;
        const projectTypes = projectDetails.projectTypes;
        const targetProjectVisibleCount = projectsTargetIndicatorOnly
            ? projectDetails.totalCount || 0
            : 0;
        const projectsControlsHtml = renderProjectsControls(targetProjectVisibleCount, {
            showTargetIndicatorFilter: targetIndicatorFilterAvailable
        });
        if (targetFilterSlot) targetFilterSlot.innerHTML = projectsControlsHtml;
        
        let html = `
            ${targetFilterSlot ? '' : projectsControlsHtml}
            <div class="projects-container">
                <div class="projects-header-row">
                    <div class="projects-header-cell projects-header-title">Статус / Вид проекта</div>
                    ${projectTypes.map(type => `<div class="projects-header-cell">${escapeHtml(type)}</div>`).join('')}
                </div>
        `;
        
        statuses.forEach((status, statusIndex) => {
            html += `
                <div class="projects-row">
                    <div class="projects-cell projects-row-title">${escapeHtml(status)}</div>
            `;
            
            projectTypes.forEach((type, typeIndex) => {
                const compositeKey = `${status}||${type}`;
                const projects = projectsByStatusAndType[compositeKey] || [];
                const count = String(getUniqueProjectTaskCount(projects));
                const hasProjects = projects.length > 0;
                const isActiveStatus = String(status).trim().toLowerCase() === 'в работе';
                const hasCriticalProjects = isActiveStatus && hasProjects && projects.some(project =>
                    project.dateColor === INDICATOR_COLORS.RED || project.hoursColor === INDICATOR_COLORS.RED
                );
                const criticalCellClass = hasCriticalProjects ? ' projects-count-cell--critical' : '';
                const criticalCellTitle = hasCriticalProjects
                    ? ' title="Есть проблемные проекты: просрочка или перерасход"'
                    : '';
                
                html += `
                    <div class="projects-cell projects-count-cell ${hasProjects ? 'clickable' : ''}${criticalCellClass}"${criticalCellTitle}
                         data-status="${escapeHtml(status)}"
                         data-type="${escapeHtml(type)}"
                         data-status-index="${statusIndex}"
                         data-type-index="${typeIndex}"
                         data-has-projects="${hasProjects}">
                        <span class="projects-count-value">
                            <span class="projects-count-number">${escapeHtml(count)}</span>
                        </span>
                        ${hasProjects ? '<span class="toggle-icon-small">▼</span>' : ''}
                    </div>
                `;
            });
            
            html += `</div>`;
            
            projectTypes.forEach((type, typeIndex) => {
                const compositeKey = `${status}||${type}`;
                const projects = projectsByStatusAndType[compositeKey] || [];
                
                if (projects.length > 0) {
                    const detailsId = `details-${statusIndex}-${typeIndex}`;
                    html += `
                        <div class="projects-details-wrapper" 
                             id="${detailsId}"
                             data-status="${escapeHtml(status)}"
                             data-type="${escapeHtml(type)}"
                             style="display: none;">
                            <div class="projects-details-content">
                                ${renderProjectDetails(projects, status, type)}
                            </div>
                        </div>
                    `;
                }
            });
        });
        
        html += `
            </div>
            ${renderProjectTableDiagnostics(projectDetails.diagnostics)}
        `;
        
        const container = document.getElementById('projects-container');
        if (container) {
            container.innerHTML = html;
            attachProjectsGroupingModeHandler();
            attachProjectsTargetIndicatorFilterHandler();
            attachProjectsOverdueFilterHandler();
            attachProjectClickHandlers();
            restoreExpandedProjectDetails();
        }
        
    } catch (error) {
        console.error('Ошибка в updateProjects:', error);
    }
}

function renderProjectsControls(visibleCount = 0, options = {}) {
    const showTargetIndicatorFilter = options.showTargetIndicatorFilter !== false;

    return `
        <div class="projects-controls">
            ${renderProjectsGroupingModeToggle()}
            ${showTargetIndicatorFilter ? renderProjectsTargetIndicatorFilter(visibleCount) : ''}
            ${renderProjectsOverdueFilter()}
        </div>
    `;
}

function isProjectsTargetIndicatorFilterAvailable() {
    const center = normalizeProjectCenterLabel(currentSheet);
    return !center || !TARGET_INDICATOR_EXCLUDED_CENTERS.has(center);
}

function renderProjectsGroupingModeToggle() {
    const firstLabel = currentSheet === 'ТехБлок' ? 'По центрам' : 'По управлениям';

    return `
        <div class="projects-grouping-toggle projects-segment-toggle" role="group" aria-label="Режим группировки проектов">
            <button type="button"
                    class="projects-grouping-toggle__button projects-segment-toggle__button${projectsGroupedByTask ? '' : ' is-active'}"
                    data-projects-grouping-mode="management"
                    aria-pressed="${projectsGroupedByTask ? 'false' : 'true'}">
                ${escapeHtml(firstLabel)}
            </button>
            <button type="button"
                    class="projects-grouping-toggle__button projects-segment-toggle__button${projectsGroupedByTask ? ' is-active' : ''}"
                    data-projects-grouping-mode="task"
                    aria-pressed="${projectsGroupedByTask ? 'true' : 'false'}">
                По задачам
            </button>
        </div>
    `;
}

function renderProjectsTargetIndicatorFilter(visibleCount = 0) {
    const upcTitle = projectsTargetIndicatorOnly
        ? `Показаны задачи УПЦ: ${visibleCount}`
        : 'Показать только задачи из расчета целевого показателя для выбранного полугодия';

    return `
        <div class="projects-target-filter projects-segment-toggle" role="group" aria-label="Фильтр задач УПЦ">
            <button type="button"
                    class="projects-segment-toggle__button${projectsTargetIndicatorOnly ? '' : ' is-active'}"
                    data-projects-target-indicator-mode="all"
                    aria-pressed="${projectsTargetIndicatorOnly ? 'false' : 'true'}">
                Все задачи
            </button>
            <button type="button"
                    class="projects-segment-toggle__button${projectsTargetIndicatorOnly ? ' is-active' : ''}"
                    data-projects-target-indicator-mode="upc"
                    aria-pressed="${projectsTargetIndicatorOnly ? 'true' : 'false'}"
                    title="${escapeHtml(upcTitle)}">
                Задачи УПЦ
            </button>
        </div>
    `;
}

function renderProjectsOverdueFilter() {
    const options = [
        { value: 'all', label: 'Все сроки' },
        { value: 'date', label: 'Просрочка даты' },
        { value: 'hours', label: 'Просрочка часов' },
        { value: 'date-hours', label: 'Дата или часы' },
        { value: 'no-date', label: 'Без даты завершения' },
        { value: 'no-hours', label: 'Без плановых часов' }
    ];

    return `
        <div class="projects-overdue-filter projects-segment-toggle" role="group" aria-label="Фильтр просрочки проектов">
            ${options.map(option => `
                <button type="button"
                        class="projects-segment-toggle__button${projectsOverdueFilter === option.value ? ' is-active' : ''}"
                        data-projects-overdue-filter="${escapeHtml(option.value)}"
                        aria-pressed="${projectsOverdueFilter === option.value ? 'true' : 'false'}">
                    ${escapeHtml(option.label)}
                </button>
            `).join('')}
        </div>
    `;
}

function attachProjectsGroupingModeHandler() {
    document.querySelectorAll('[data-projects-grouping-mode]').forEach(button => {
        button.addEventListener('click', () => {
            const nextGroupedByTask = button.dataset.projectsGroupingMode === 'task';
            if (projectsGroupedByTask === nextGroupedByTask) return;

            projectsGroupedByTask = nextGroupedByTask;
            updateProjects();
        });
    });
}

function attachProjectsTargetIndicatorFilterHandler() {
    document.querySelectorAll('[data-projects-target-indicator-mode]').forEach(button => {
        button.addEventListener('click', () => {
            const nextTargetOnly = button.dataset.projectsTargetIndicatorMode === 'upc';
            if (projectsTargetIndicatorOnly === nextTargetOnly) return;

            projectsTargetIndicatorOnly = nextTargetOnly;
            updateProjects();
        });
    });
}

function attachProjectsOverdueFilterHandler() {
    document.querySelectorAll('[data-projects-overdue-filter]').forEach(button => {
        button.addEventListener('click', () => {
            const nextFilter = button.dataset.projectsOverdueFilter || 'all';
            if (projectsOverdueFilter === nextFilter) return;

            projectsOverdueFilter = nextFilter;
            updateProjects();
        });
    });
}

function attachProjectClickHandlers() {
    document.querySelectorAll('.projects-count-cell.clickable').forEach(cell => {
        cell.addEventListener('click', handleProjectCellClick);
    });
}

function getProjectDetailsExpandKey(status, type) {
    return `${status || ''}||${type || ''}`;
}

function restoreExpandedProjectDetails() {
    if (!expandedProjectDetailsKey) return;

    const targetCell = Array.from(document.querySelectorAll('.projects-count-cell.clickable')).find(cell =>
        getProjectDetailsExpandKey(cell.dataset.status, cell.dataset.type) === expandedProjectDetailsKey
    );
    if (!targetCell) {
        expandedProjectDetailsKey = null;
        return;
    }

    const targetWrapper = Array.from(document.querySelectorAll('.projects-details-wrapper')).find(wrapper =>
        getProjectDetailsExpandKey(wrapper.dataset.status, wrapper.dataset.type) === expandedProjectDetailsKey
    );
    if (!targetWrapper) {
        expandedProjectDetailsKey = null;
        return;
    }

    targetWrapper.style.display = 'block';
    targetCell.classList.add('expanded');
    const icon = targetCell.querySelector('.toggle-icon-small');
    if (icon) icon.textContent = '▲';
}

function handleProjectCellClick(event) {
    const cell = event.currentTarget;
    const status = cell.dataset.status;
    const type = cell.dataset.type;
    
    const allWrappers = document.querySelectorAll('.projects-details-wrapper');
    let targetWrapper = null;
    
    for (let wrapper of allWrappers) {
        if (wrapper.dataset.status === status && wrapper.dataset.type === type) {
            targetWrapper = wrapper;
            break;
        }
    }
    
    if (!targetWrapper) {
        console.warn(`Wrapper не найден для: ${status} / ${type}`);
        return;
    }
    
    const isCurrentlyOpen = targetWrapper.style.display === 'block';
    
    allWrappers.forEach(wrapper => {
        wrapper.style.display = 'none';
    });
    
    document.querySelectorAll('.projects-count-cell').forEach(c => {
        c.classList.remove('expanded');
        const icon = c.querySelector('.toggle-icon-small');
        if (icon) icon.textContent = '▼';
    });
    
    if (!isCurrentlyOpen) {
        expandedProjectDetailsKey = getProjectDetailsExpandKey(status, type);
        targetWrapper.style.display = 'block';
        cell.classList.add('expanded');
        const icon = cell.querySelector('.toggle-icon-small');
        if (icon) icon.textContent = '▲';
    } else {
        expandedProjectDetailsKey = null;
    }
}

function getProjectsFromDetails(options = {}) {
    const targetProjectKeys = options.targetProjectKeys instanceof Set ? options.targetProjectKeys : null;
    const currentCenter = normalizeProjectCenterLabel(currentSheet);
    const includeAllCenters = currentSheet === 'ТехБлок';
    const projects = getProjectsFromTechBlockForCenter(currentCenter, {
        targetProjectKeys,
        groupByTask: false,
        includeAllCenters
    });
    const filteredProjects = filterProjectTableByOverdue(projects, projectsOverdueFilter);
    const projectsByKey = getProjectTableProjectsByKey(filteredProjects);

    return {
        projectsByKey,
        diagnostics: getProjectTableDiagnostics({
            centerKey: currentCenter,
            includeAllCenters,
            targetProjectKeys
        }),
        totalCount: getUniqueProjectTaskCount(filteredProjects),
        statuses: getOrderedProjectTableStatuses(filteredProjects),
        projectTypes: getOrderedProjectTableTypes(filteredProjects)
    };
}

const PROJECT_TABLE_STATUS_ORDER = ['В работе', 'На паузе', 'В очереди', 'У заказчика', 'Завершена'];
const PROJECT_TABLE_ALLOWED_STATUSES = new Set(PROJECT_TABLE_STATUS_ORDER);
const PROJECT_TABLE_DIAGNOSTIC_IGNORED_IDS = new Set(['1001', '1002', '-']);

function shouldIgnoreProjectTableRowByStatus(status) {
    const normalizedStatus = normalizeProjectStage(status);
    return normalizedStatus.type === 'cancelled'
        || !PROJECT_TABLE_ALLOWED_STATUSES.has(normalizedStatus.label);
}

function normalizeProjectTableDiagnosticRawId(value) {
    if (!isValid(value)) return '';
    if (typeof value === 'number' && Number.isFinite(value)) {
        return Number.isInteger(value) ? String(value) : String(value).replace(/\.0$/, '');
    }

    return String(value).trim().replace(/\.0$/, '');
}

function shouldIgnoreProjectTableDiagnosticRow(rawId, rawStatus, rawProjectType) {
    const id = normalizeProjectTableDiagnosticRawId(rawId);
    if (!PROJECT_TABLE_DIAGNOSTIC_IGNORED_IDS.has(id)) return false;

    return !isFilledProjectStageValue(rawStatus) && !getCleanTextValue(rawProjectType);
}

function getProjectTableProjectsByKey(projects) {
    const groupedProjects = {};
    if (!Array.isArray(projects)) return groupedProjects;

    projects.forEach(project => {
        const statusKey = project.status || 'Без статуса';
        const typeKey = project.type || 'Другие проекты';
        const compositeKey = `${statusKey}||${typeKey}`;

        if (!groupedProjects[compositeKey]) {
            groupedProjects[compositeKey] = [];
        }

        groupedProjects[compositeKey].push(project);
    });

    return groupedProjects;
}

function getProjectTableStableCountsByKey(projects) {
    const groupedProjects = getProjectTableProjectsByKey(projects);

    return Object.fromEntries(
        Object.entries(groupedProjects).map(([key, items]) => [key, getUniqueProjectTaskCount(items)])
    );
}

function getProjectTableDiagnostics(options = {}) {
    if (!workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) return [];

    const centerKey = options.centerKey || '';
    const includeAllCenters = options.includeAllCenters === true;
    const targetProjectKeys = options.targetProjectKeys instanceof Set ? options.targetProjectKeys : null;
    const rows = getWorksheetRowsByHeaders('Проекты_ТехБлок');
    const projectNamesLookup = getProjectNamesLookup();
    const rowGroups = new Map();
    const diagnostics = [];

    rows.forEach((row, index) => {
        const rawId = getWorksheetRowCellValue(row, 1);
        const rawProjectType = getRowValue(row, ['Вид проекта']);
        const rawStatus = getRowValue(row, ['Статус']);
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const id = normalizeProjectId(rawId);
        const projectType = getCleanTextValue(rawProjectType) || 'Другие проекты';
        const status = normalizeProjectStage(rawStatus);
        const management = normalizeManagementLabel(getRowValue(row, ['Управление'])) || 'Без управления';
        const displayName = projectNamesLookup.get(normalizeProjectIdForLookup(rawId));
        const name = displayName || getCleanTextValue(getRowValue(row, ['Наименование проекта', 'Проект'])) || 'Без названия';

        if (shouldIgnoreProjectTableRowByStatus(rawStatus)) return;
        if (shouldIgnoreProjectTableDiagnosticRow(rawId, rawStatus, rawProjectType)) return;

        if (!center || !id) {
            diagnostics.push({
                center: center || 'Без центра',
                id: id || `строка ${index + 2}`,
                name,
                problem: !center && !id ? 'не заполнены центр и ID' : !center ? 'не заполнен центр' : 'не заполнен ID'
            });
            return;
        }

        if (!includeAllCenters && center !== centerKey) return;

        const targetKey = getTargetIndicatorProjectKey(center, id);
        if (targetProjectKeys && !targetProjectKeys.has(targetKey)) return;
        if (targetProjectKeys && !isTargetIndicatorProjectTypeIncluded(projectType)) return;

        const groupKey = `${center}||${id}`;
        if (!rowGroups.has(groupKey)) {
            rowGroups.set(groupKey, {
                center,
                id,
                names: new Set(),
                types: new Set(),
                statuses: new Set(),
                managements: new Set()
            });
        }

        const group = rowGroups.get(groupKey);
        group.names.add(name);
        group.types.add(projectType);
        group.statuses.add(status.label || 'Без статуса');
        group.managements.add(management);
    });

    rowGroups.forEach(group => {
        const problems = [];
        const types = Array.from(group.types).filter(Boolean);

        if (types.length > 1) {
            problems.push(`несколько видов проекта: ${types.join(', ')}`);
        }

        if (problems.length === 0) return;

        diagnostics.push({
            center: group.center,
            id: group.id,
            name: Array.from(group.names).filter(Boolean).slice(0, 3).join(', ') || 'Без названия',
            problem: problems.join('; ')
        });
    });

    return diagnostics.sort((a, b) => {
        const centerDiff = compareProjectTableCenters(a.center, b.center);
        if (centerDiff !== 0) return centerDiff;
        return String(a.id).localeCompare(String(b.id), 'ru', { numeric: true });
    });
}

function renderProjectTableDiagnostics(diagnostics) {
    if (!Array.isArray(diagnostics) || diagnostics.length === 0) return '';

    return `
        <details class="project-diagnostics">
            <summary class="project-diagnostics__summary">
                <span>Задачи для проверки</span>
                <small>${escapeHtml(String(diagnostics.length))}</small>
            </summary>
            <div class="project-diagnostics__body">
                <div class="project-diagnostics__intro">Эти задачи имеют неоднозначные строки в источнике или строки, которые не попадают в общую матрицу проектов.</div>
                <div class="project-diagnostics__list">
                    ${diagnostics.map(item => `
                        <div class="project-diagnostics__row">
                            <div class="project-diagnostics__id">${escapeHtml(item.center)} · ${escapeHtml(item.id)}</div>
                            <div class="project-diagnostics__name">${escapeHtml(item.name)}</div>
                            <div class="project-diagnostics__problem">${escapeHtml(item.problem)}</div>
                        </div>
                    `).join('')}
                </div>
            </div>
        </details>
    `;
}

function getProjectsFromTechBlockForCenter(centerKey, options = {}) {
    const targetProjectKeys = options.targetProjectKeys instanceof Set ? options.targetProjectKeys : null;
    const groupByTask = Boolean(options.groupByTask);
    const includeAllCenters = options.includeAllCenters === true;
    if ((!centerKey && !includeAllCenters) || !workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) return [];

    const rows = getWorksheetRowsByHeaders('Проекты_ТехБлок');
    const projectNamesLookup = getProjectNamesLookup();
    const projectMap = new Map();

    rows.forEach(row => {
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const id = normalizeProjectIdForLookup(getWorksheetRowCellValue(row, 1));
        if (!center || (!includeAllCenters && center !== centerKey) || !id) return;

        const status = getRowValue(row, ['Статус']);
        if (shouldIgnoreProjectTableRowByStatus(status)) return;

        const targetKey = getTargetIndicatorProjectKey(center, id);
        if (targetProjectKeys && !targetProjectKeys.has(targetKey)) return;

        const management = normalizeManagementLabel(getRowValue(row, ['Управление']));
        const projectType = getCleanTextValue(getRowValue(row, ['Вид проекта'])) || 'Другие проекты';
        if (targetProjectKeys && !isTargetIndicatorProjectTypeIncluded(projectType)) return;

        const statusLabel = normalizeProjectStage(status).label || 'Без статуса';
        const key = getProjectTableGroupKey(center, id, management, projectType, statusLabel, { groupByTask, includeAllCenters });

        if (!projectMap.has(key)) {
            projectMap.set(key, {
                key,
                targetKey,
                center,
                centers: new Set(),
                id,
                name: '',
                managements: new Set(),
                type: '',
                statuses: [],
                planDate: null,
                estimateDate: null,
                planHours: 0,
                actualHours: 0,
                factStart: null,
                factEnd: null,
                hasComment: false,
                comment: ''
            });
        }

        const project = projectMap.get(key);
        const name = getCleanTextValue(getRowValue(row, ['Наименование проекта', 'Проект']));
        const planHours = parseNumericValue(getRowValue(row, ['Количество ПЛАНовых часов', 'Плановые часы']));
        const actualHours = parseNumericValue(getRowValue(row, ['Количество ФАКТических часов', 'Фактические часы']));
        const rowComment = getCleanTextValue(getRowValue(row, ['Комментарий']));

        if (!project.name && name) project.name = name;
        if (center) project.centers.add(center);
        if (management) project.managements.add(management);
        if (!project.type && projectType) project.type = projectType;
        if (rowComment) {
            project.hasComment = true;
            project.comment = project.comment ? `${project.comment}\n${rowComment}` : rowComment;
        }
        project.statuses.push(status);
        project.planDate = getLaterDate(project.planDate, parseDateValue(getRowValue(row, [
            'План дата завершения этапа по графику проекта',
            'Дата завершения работ',
            'Дата завершения работ по договору',
            'Плановая дата окончания'
        ])));
        project.estimateDate = getLaterDate(project.estimateDate, parseDateValue(getRowValue(row, [
            'ПРОГНОЗ окончания',
            'Прогноз окончания',
            'Прогнозируемая дата окончания'
        ])));
        project.factStart = getEarlierDate(project.factStart, parseDateValue(getRowValue(row, [
            'ФАКТ начало',
            'Фактическая дата начала'
        ])));
        project.factEnd = getLaterDate(project.factEnd, parseDateValue(getRowValue(row, [
            'ФАКТ окончание',
            'Фактическая дата окончания'
        ])));
        if (planHours !== null) project.planHours += planHours;
        if (actualHours !== null) project.actualHours += actualHours;
    });

    return Array.from(projectMap.values()).map(project => {
        const status = getProjectTableAggregatedStatus(project.statuses);
        const remainingHours = project.planHours - project.actualHours;
        const managements = Array.from(project.managements)
            .filter(Boolean)
            .sort((a, b) => a.localeCompare(b, 'ru', { numeric: true }));
        const centers = Array.from(project.centers)
            .filter(Boolean)
            .sort(compareProjectTableCenters);
        const displayName = projectNamesLookup.get(normalizeProjectIdForLookup(project.id));

        return {
            targetKey: project.targetKey,
            countKey: getTargetIndicatorProjectKey(project.center, project.id),
            taskKey: normalizeProjectId(project.id),
            center: project.center,
            centers,
            centerLabel: centers.length > 0 ? centers.join(', ') : project.center,
            id: project.id,
            name: project.name,
            displayName: displayName || project.name,
            management: managements.join(', ') || 'Без управления',
            managements,
            type: project.type || 'Другие проекты',
            status: status.label,
            statusType: status.type,
            planDate: project.planDate,
            estimateDate: project.estimateDate,
            factStart: project.factStart,
            factEnd: project.factEnd,
            dateIndicator: getProjectTableDateIndicator(project.planDate, project.estimateDate),
            dateColor: getProjectTableDateColor(project.planDate, project.estimateDate),
            doneDateIndicator: getProjectTableDateIndicator(project.planDate, project.factEnd),
            doneDateColor: getProjectTableDateColor(project.planDate, project.factEnd),
            planHours: project.planHours,
            actualHours: project.actualHours,
            remainingHours,
            hoursColor: getProjectTableHoursColor(remainingHours),
            hasComment: project.hasComment === true,
            comment: project.comment || ''
        };
    }).filter(project => PROJECT_TABLE_ALLOWED_STATUSES.has(project.status)).sort((a, b) => {
        if (includeAllCenters) {
            if (!groupByTask) {
                const centerDiff = compareProjectTableCenters(a.centers?.[0] || a.center, b.centers?.[0] || b.center);
                if (centerDiff !== 0) return centerDiff;
            }

            return String(a.id).localeCompare(String(b.id), 'ru', { numeric: true });
        }

        const managementDiff = a.management.localeCompare(b.management, 'ru', { numeric: true });
        if (managementDiff !== 0) return managementDiff;
        return String(a.id).localeCompare(String(b.id), 'ru', { numeric: true });
    });
}

function getUniqueProjectTaskCount(projects) {
    if (!Array.isArray(projects) || projects.length === 0) return 0;

    const taskKeys = new Set();
    projects.forEach((project, index) => {
        const key = project?.countKey
            || getTargetIndicatorProjectKey(project?.center, project?.id)
            || project?.taskKey
            || normalizeProjectId(project?.id)
            || project?.targetKey
            || project?.key
            || `project-${index}`;

        taskKeys.add(key);
    });

    return taskKeys.size;
}

function compareProjectTableCenters(a, b) {
    const order = ['ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ', 'ЛОИ', 'УВП'];
    const aCenter = normalizeProjectCenterLabel(a) || String(a || '');
    const bCenter = normalizeProjectCenterLabel(b) || String(b || '');
    const aIndex = order.indexOf(aCenter);
    const bIndex = order.indexOf(bCenter);
    const aRank = aIndex === -1 ? order.length : aIndex;
    const bRank = bIndex === -1 ? order.length : bIndex;

    if (aRank !== bRank) return aRank - bRank;
    return aCenter.localeCompare(bCenter, 'ru', { numeric: true });
}

function getProjectTableGroupKey(center, id, management, projectType, statusLabel, options = {}) {
    if (options.includeAllCenters) {
        if (options.groupByTask) {
            return normalizeProjectId(id);
        }

        return [
            normalizeProjectCenterLabel(center),
            normalizeProjectId(id),
            normalizeLookupText(management) || 'без управления',
            normalizeLookupText(projectType) || 'другие проекты',
            normalizeLookupText(statusLabel) || 'без статуса'
        ].join('||');
    }

    if (options.groupByTask) {
        return [
            normalizeProjectCenterLabel(center),
            normalizeProjectId(id)
        ].join('||');
    }

    return [
        normalizeProjectCenterLabel(center),
        normalizeProjectId(id),
        normalizeLookupText(management) || 'без управления',
        normalizeLookupText(projectType) || 'другие проекты',
        normalizeLookupText(statusLabel) || 'без статуса'
    ].join('||');
}

function getProjectTableAggregatedStatus(statuses) {
    const normalizedStatuses = statuses.map(status => normalizeProjectStage(status));
    const hasDone = normalizedStatuses.some(status => status.type === 'done');
    const hasNotDone = normalizedStatuses.some(status => status.type !== 'done');

    if (hasDone && hasNotDone) {
        return { label: 'В работе', type: 'active' };
    }

    return getOverallProjectStatus(statuses);
}

function getProjectTableDateIndicator(planDate, estimateDate) {
    if (!(planDate instanceof Date) || isNaN(planDate.getTime()) || !(estimateDate instanceof Date) || isNaN(estimateDate.getTime())) {
        return '';
    }

    if (estimateDate <= planDate) {
        return 'Окончание в срок';
    }

    const delayDays = Math.max(1, Math.ceil((estimateDate.getTime() - planDate.getTime()) / 86400000));
    return `Отставание на ${delayDays} ${getRussianDayWord(delayDays)}`;
}

function getRussianDayWord(value) {
    const absValue = Math.abs(Number(value));
    const lastTwoDigits = absValue % 100;
    const lastDigit = absValue % 10;

    if (lastTwoDigits >= 11 && lastTwoDigits <= 14) return 'дней';
    if (lastDigit === 1) return 'день';
    if (lastDigit >= 2 && lastDigit <= 4) return 'дня';
    return 'дней';
}

function getProjectTableDateColor(planDate, estimateDate) {
    if (!(planDate instanceof Date) || isNaN(planDate.getTime()) || !(estimateDate instanceof Date) || isNaN(estimateDate.getTime())) {
        return INDICATOR_COLORS.GRAY;
    }

    return estimateDate <= planDate ? INDICATOR_COLORS.GREEN : INDICATOR_COLORS.RED;
}

function getProjectTableHoursColor(remainingHours) {
    if (!Number.isFinite(remainingHours)) return INDICATOR_COLORS.GRAY;
    if (remainingHours < 0) return INDICATOR_COLORS.RED;
    if (remainingHours < 24.5) return INDICATOR_COLORS.YELLOW;
    return INDICATOR_COLORS.GREEN;
}

function aggregateProjectDetailRowsByTask(projects) {
    if (!Array.isArray(projects) || projects.length === 0) return [];

    const projectMap = new Map();

    projects.forEach((project, index) => {
        const taskKey = project?.taskKey || normalizeProjectId(project?.id) || project?.key || `project-${index}`;
        if (!projectMap.has(taskKey)) {
            projectMap.set(taskKey, {
                targetKey: project.targetKey,
                countKey: project.countKey,
                taskKey,
                center: project.center,
                centers: new Set(Array.isArray(project.centers) ? project.centers : [project.center].filter(Boolean)),
                id: project.id,
                name: project.name || '',
                displayName: project.displayName || project.name || '',
                managements: new Set(Array.isArray(project.managements) ? project.managements : [project.management].filter(Boolean)),
                type: project.type || 'Другие проекты',
                status: project.status || 'Без статуса',
                statusType: project.statusType || normalizeProjectStage(project.status).type,
                planDate: project.planDate || null,
                estimateDate: project.estimateDate || null,
                factStart: project.factStart || null,
                factEnd: project.factEnd || null,
                planHours: 0,
                actualHours: 0,
                hasComment: false,
                comment: ''
            });
        }

        const current = projectMap.get(taskKey);
        if (!current.name && project.name) current.name = project.name;
        if (!current.displayName && project.displayName) current.displayName = project.displayName;
        if (!current.type && project.type) current.type = project.type;
        if (!current.status && project.status) current.status = project.status;
        if (!current.statusType && project.statusType) current.statusType = project.statusType;
        if (project.hasComment) current.hasComment = true;
        if (project.comment) {
            current.comment = current.comment ? `${current.comment}\n${project.comment}` : project.comment;
        }

        (Array.isArray(project.centers) ? project.centers : [project.center]).filter(Boolean).forEach(center => current.centers.add(center));
        (Array.isArray(project.managements) ? project.managements : [project.management]).filter(Boolean).forEach(management => current.managements.add(management));

        current.planDate = getLaterDate(current.planDate, project.planDate);
        current.estimateDate = getLaterDate(current.estimateDate, project.estimateDate);
        current.factStart = getEarlierDate(current.factStart, project.factStart);
        current.factEnd = getLaterDate(current.factEnd, project.factEnd);
        current.planHours += Number(project.planHours) || 0;
        current.actualHours += Number(project.actualHours) || 0;
    });

    return Array.from(projectMap.values()).map(project => {
        const centers = Array.from(project.centers).filter(Boolean).sort(compareProjectTableCenters);
        const managements = Array.from(project.managements).filter(Boolean).sort((a, b) => a.localeCompare(b, 'ru', { numeric: true }));
        const remainingHours = project.planHours - project.actualHours;

        return {
            targetKey: project.targetKey,
            countKey: project.countKey,
            taskKey: project.taskKey,
            center: project.center,
            centers,
            centerLabel: centers.join(', ') || project.center,
            id: project.id,
            name: project.name,
            displayName: project.displayName || project.name,
            management: managements.join(', ') || 'Без управления',
            managements,
            type: project.type || 'Другие проекты',
            status: project.status || 'Без статуса',
            statusType: project.statusType,
            planDate: project.planDate,
            estimateDate: project.estimateDate,
            factStart: project.factStart,
            factEnd: project.factEnd,
            dateIndicator: getProjectTableDateIndicator(project.planDate, project.estimateDate),
            dateColor: getProjectTableDateColor(project.planDate, project.estimateDate),
            doneDateIndicator: getProjectTableDateIndicator(project.planDate, project.factEnd),
            doneDateColor: getProjectTableDateColor(project.planDate, project.factEnd),
            planHours: project.planHours,
            actualHours: project.actualHours,
            remainingHours,
            hoursColor: getProjectTableHoursColor(remainingHours),
            hasComment: project.hasComment === true,
            comment: project.comment || ''
        };
    }).sort((a, b) => String(a.id).localeCompare(String(b.id), 'ru', { numeric: true }));
}

function getProjectTableBadgeTone(color) {
    if (color === INDICATOR_COLORS.RED) return 'danger';
    if (color === INDICATOR_COLORS.YELLOW) return 'warning';
    if (color === INDICATOR_COLORS.GREEN) return 'success';
    return 'neutral';
}

function filterProjectTableByOverdue(projects, filter) {
    if (!Array.isArray(projects) || filter === 'all') return projects;

    return projects.filter(project => {
        const hasDateOverdue = isProjectTableDateOverdue(project);
        const hasHoursOverdue = isProjectTableHoursOverdue(project);

        if (filter === 'date') return hasDateOverdue;
        if (filter === 'hours') return hasHoursOverdue;
        if (filter === 'date-hours') return hasDateOverdue || hasHoursOverdue;
        if (filter === 'no-date') return isProjectTableMissingPlanDate(project);
        if (filter === 'no-hours') return isProjectTableMissingPlanHours(project);
        return true;
    });
}

function isProjectTableDateOverdue(project) {
    if (!project) return false;
    if (project.statusType === 'done') {
        return project.doneDateColor === INDICATOR_COLORS.RED;
    }
    if (getProjectTableFactEndDelayDays(project) > 0) return true;
    return project.dateColor === INDICATOR_COLORS.RED;
}

function isProjectTableHoursOverdue(project) {
    return Boolean(project) && project.hoursColor === INDICATOR_COLORS.RED;
}

function isProjectTableMissingPlanDate(project) {
    return !project || !(project.planDate instanceof Date) || isNaN(project.planDate.getTime());
}

function isProjectTableMissingPlanHours(project) {
    return Boolean(project) && !(Number(project.planHours) > 0);
}

function getOrderedProjectTableStatuses(projects) {
    return PROJECT_TABLE_STATUS_ORDER.filter(status => projects.some(project => project.status === status));
}

function getOrderedProjectTableTypes(projects) {
    const order = ['Коммерческий проект', 'Инвестиционный проект', 'Предпроектная подготовка', 'РИД', 'Другие проекты'];
    return orderProjectTableValues(Array.from(new Set(projects.map(project => project.type).filter(Boolean))), order);
}

function orderProjectTableValues(values, order) {
    return values.sort((a, b) => {
        const aIndex = order.indexOf(a);
        const bIndex = order.indexOf(b);
        const aRank = aIndex === -1 ? order.length : aIndex;
        const bRank = bIndex === -1 ? order.length : bIndex;
        if (aRank !== bRank) return aRank - bRank;
        return a.localeCompare(b, 'ru', { numeric: true });
    });
}

function getDateIndicatorColor(colorValue, indicatorText) {
    if (isValid(indicatorText)) {
        const text = String(indicatorText).toLowerCase().trim();
        if (text === 'окончание в срок') {
            return INDICATOR_COLORS.GREEN;
        }
    }
    
    if (!isValid(colorValue)) return INDICATOR_COLORS.GRAY;
    
    const numValue = parseFloat(colorValue);
    if (isNaN(numValue)) return INDICATOR_COLORS.GRAY;
    
    if (numValue >= 0) return INDICATOR_COLORS.GREEN;
    if (numValue >= -5) return INDICATOR_COLORS.YELLOW;
    return INDICATOR_COLORS.RED;
}

function getHoursIndicatorColor(value) {
    if (!isValid(value)) return INDICATOR_COLORS.GRAY;
    
    const numValue = parseFloat(value);
    if (isNaN(numValue)) return INDICATOR_COLORS.GRAY;
    
    if (numValue < 0) return INDICATOR_COLORS.RED;
    if (numValue < 24.5) return INDICATOR_COLORS.YELLOW;
    return INDICATOR_COLORS.GREEN;
}

function shouldGroupProjectDetailsByManagement() {
    if (currentSheet === 'ТехБлок') return false;
    if (projectsGroupedByTask) return false;

    const center = normalizeProjectCenterLabel(currentSheet);
    return center !== 'ЦОТ' && center !== 'ЦДПМ';
}

function ensureProjectCommentPopover() {
    let popover = document.getElementById('project-comment-popover');
    if (!popover) {
        document.body.insertAdjacentHTML('beforeend', '<div id="project-comment-popover" class="project-comment-popover" hidden></div>');
        popover = document.getElementById('project-comment-popover');
    }
    return popover;
}

let projectCommentPopoverAnchor = null;

function openProjectCommentPopover(anchorEl, commentText) {
    const popover = ensureProjectCommentPopover();

    if (!popover.hidden && projectCommentPopoverAnchor === anchorEl) {
        closeProjectCommentPopover();
        return;
    }

    popover.innerHTML = `
        <div class="project-comment-popover__header">
            <span class="project-comment-popover__title">Комментарий</span>
            <button type="button" class="project-comment-popover__close" aria-label="Закрыть">✖</button>
        </div>
        <div class="project-comment-popover__text">${escapeHtml(commentText)}</div>
    `;

    projectCommentPopoverAnchor = anchorEl;
    positionPopover(popover, anchorEl);
    popover.hidden = false;

    popover.querySelector('.project-comment-popover__close').addEventListener('click', closeProjectCommentPopover);
}

function closeProjectCommentPopover() {
    const popover = document.getElementById('project-comment-popover');
    if (popover) {
        popover.hidden = true;
        popover.innerHTML = '';
    }
    projectCommentPopoverAnchor = null;
}

document.addEventListener('click', (e) => {
    const flag = e.target.closest('.project-comment-flag');
    if (flag) {
        e.stopPropagation();
        openProjectCommentPopover(flag, flag.dataset.comment || '');
        return;
    }

    const popover = document.getElementById('project-comment-popover');
    if (!popover || popover.hidden) return;
    if (popover.contains(e.target)) return;
    closeProjectCommentPopover();
});
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeProjectCommentPopover();
});

function renderProjectDetails(projects, status, type) {
    if (!projects || projects.length === 0) {
        return '<div class="empty-state-note">Нет данных о проектах</div>';
    }

    const detailProjects = projectsGroupedByTask
        ? aggregateProjectDetailRowsByTask(projects)
        : projects;
    const normalizedStatus = normalizeProjectStage(status);
    const isDoneStatus = normalizedStatus.type === 'done';
    const showDateIndicatorColumn = normalizedStatus.type === 'active';
    const isTechBlockProjects = currentSheet === 'ТехБлок';
    const groupTechBlockByCenter = isTechBlockProjects && !projectsGroupedByTask;
    const showCenterColumn = isTechBlockProjects && projectsGroupedByTask;
    const showManagementColumn = projectsGroupedByTask && !isTechBlockProjects;
    const showContextColumn = showCenterColumn || showManagementColumn;
    const groupByManagement = shouldGroupProjectDetailsByManagement();
    
    const detailGroups = new Map();
    if (groupTechBlockByCenter) {
        detailProjects.forEach(project => {
            getProjectDetailCenterGroups(project).forEach(center => {
                if (!detailGroups.has(center)) detailGroups.set(center, []);
                detailGroups.get(center).push(project);
            });
        });
    } else if (groupByManagement) {
        detailProjects.forEach(project => {
            getProjectDetailManagementGroups(project).forEach(management => {
                if (!detailGroups.has(management)) detailGroups.set(management, []);
                detailGroups.get(management).push(project);
            });
        });
    } else {
        detailGroups.set('all', detailProjects);
    }
    
    let html = `
        <div class="management-projects">
            <div class="section-title">Проекты: ${escapeHtml(type)} / ${escapeHtml(status)}</div>
    `;
    
    const detailGroupKeys = groupTechBlockByCenter
        ? Array.from(detailGroups.keys()).sort(compareProjectTableCenters)
        : groupByManagement
            ? Array.from(detailGroups.keys()).sort((a, b) => a.localeCompare(b, 'ru'))
            : ['all'];

    detailGroupKeys.forEach(groupKey => {
        const mgmtProjects = detailGroups.get(groupKey) || [];
        const groupTitle = groupTechBlockByCenter
            ? getProjectStageFullLabel(groupKey)
            : groupByManagement
                ? groupKey
                : '';

        html += `
            <div class="management-group">
                ${groupTitle ? `<div class="management-title">${escapeHtml(groupTitle)}</div>` : ''}
                
                <div class="projects-grid-header ${showDateIndicatorColumn || isDoneStatus ? '' : 'projects-grid-header--no-date-indicator'}${isDoneStatus ? ' projects-grid-header--done-with-indicator' : ''}${showContextColumn ? ' projects-grid-header--with-management' : ''}">
                    ${showContextColumn ? `<div class="grid-cell grid-header">${showCenterColumn ? 'Центр' : 'Управление'}</div>` : ''}
                    <div class="grid-cell grid-header">ID</div>
                    <div class="grid-cell grid-header">Название</div>
                    ${isDoneStatus
                        ? `
                            <div class="grid-cell grid-header">Факт начало</div>
                            <div class="grid-cell grid-header">Факт окончание</div>
                            <div class="grid-cell grid-header">План дата завершения этапа по графику проекта</div>
                            <div class="grid-cell grid-header">Индикатор по дате</div>
                        `
                        : `
                            <div class="grid-cell grid-header">План дата завершения этапа по графику проекта</div>
                            <div class="grid-cell grid-header">Прогнозируемая дата окончания</div>
                            ${showDateIndicatorColumn ? '<div class="grid-cell grid-header">Индикатор по дате</div>' : ''}
                        `
                    }
                    <div class="grid-cell grid-header">План/Факт часы</div>
                    <div class="grid-cell grid-header">Остаток по часам</div>
                </div>
                
                <div class="projects-grid-body">
        `;
        
        mgmtProjects.forEach(project => {
            const planDate = formatExcelDate(project.planDate);
            const estimateDate = formatExcelDate(project.estimateDate);
            const factStart = formatExcelDate(project.factStart);
            const factEnd = formatExcelDate(project.factEnd);
            const hasBothProjectDates = Boolean(planDate && estimateDate);
            
            const planFormatted = Math.round(project.planHours).toLocaleString('ru-RU');
            const factFormatted = Math.round(project.actualHours).toLocaleString('ru-RU');
            const hoursDisplay = (project.planHours || project.actualHours) 
                ? `${planFormatted} / ${factFormatted}`
                : '—';
            
            const remainingHours = Number.isFinite(project.remainingHours)
                ? Math.round(project.remainingHours)
                : null;
            const hoursIndicatorDisplay = remainingHours === null ? '—' : remainingHours.toLocaleString('ru-RU');
            const hoursBadgeTooltip = remainingHours === null
                ? ''
                : ` title="Остаток по часам: ${escapeHtml(hoursIndicatorDisplay)}"`;
            const hoursBadgeTone = getProjectTableBadgeTone(project.hoursColor);
            const hoursBadgeValueDisplay = remainingHours !== null && hoursBadgeTone === 'danger'
                ? Math.abs(remainingHours).toLocaleString('ru-RU')
                : hoursIndicatorDisplay;
            const hoursBadgeCaption = hoursBadgeTone === 'danger'
                ? 'превышение'
                : hoursBadgeTone === 'warning'
                    ? 'приближается к лимиту'
                    : '';

            const safePlanDate = escapeHtml(planDate || '-');
            const safeEstimateDate = escapeHtml(estimateDate || '-');
            const safeFactStart = escapeHtml(factStart || '-');
            const safeFactEnd = escapeHtml(factEnd || '-');
            const safeManagement = escapeHtml(project.management || 'Без управления');
            const safeCenterLabel = escapeHtml(project.centerLabel || project.center || 'Без центра');
            const displayProjectName = getCleanTextValue(project.displayName) || getCleanTextValue(project.name);
            const safeProjectName = escapeHtml(displayProjectName || '—');
            const safeHoursDisplay = escapeHtml(hoursDisplay);
            const safeHoursBadgeValueDisplay = escapeHtml(hoursBadgeValueDisplay);
            const safeHoursBadgeCaption = escapeHtml(hoursBadgeCaption);
            const factEndDelayDays = getProjectTableFactEndDelayDays(project);
            const factEndHtml = factEndDelayDays > 0
                ? `
                    <span class="project-value-badge project-value-badge--danger" title="Факт окончания позже плановой даты завершения этапа на ${escapeHtml(String(factEndDelayDays))} ${escapeHtml(getRussianDayWord(factEndDelayDays))}">${safeFactEnd}</span>
                `
                : safeFactEnd;
            const dateIndicatorHtml = hasBothProjectDates
                ? `<span class="project-value-badge project-value-badge--${escapeHtml(getProjectTableBadgeTone(project.dateColor))}">${escapeHtml(project.dateIndicator) || '-'}</span>`
                : '-';
            const hasBothDoneDates = Boolean(planDate && factEnd);
            const doneDateIndicatorHtml = hasBothDoneDates
                ? `<span class="project-value-badge project-value-badge--${escapeHtml(getProjectTableBadgeTone(project.doneDateColor))}">${escapeHtml(project.doneDateIndicator) || '-'}</span>`
                : '-';
            const commentFlagHtml = project.hasComment
                ? `<button type="button" class="project-comment-flag" data-comment="${escapeHtml(project.comment)}" title="Показать комментарий" aria-label="Показать комментарий">!</button>`
                : '';
            
            html += `
                <div class="projects-grid-row ${showDateIndicatorColumn || isDoneStatus ? '' : 'projects-grid-row--no-date-indicator'}${isDoneStatus ? ' projects-grid-row--done-with-indicator' : ''}${showContextColumn ? ' projects-grid-row--with-management' : ''}">
                    ${showContextColumn ? `<div class="grid-cell project-management" title="${showCenterColumn ? safeCenterLabel : safeManagement}">${showCenterColumn ? safeCenterLabel : safeManagement}</div>` : ''}
                    <div class="grid-cell project-id"><span class="project-id-value">${escapeHtml(project.id) || '—'}</span>${commentFlagHtml}</div>
                    <div class="grid-cell project-name">${safeProjectName}</div>
                    ${isDoneStatus
                        ? `
                            <div class="grid-cell project-date">${safeFactStart}</div>
                            <div class="grid-cell project-date">${factEndHtml}</div>
                            <div class="grid-cell project-date">${safePlanDate}</div>
                            <div class="grid-cell project-indicator-cell">${doneDateIndicatorHtml}</div>
                        `
                        : `
                            <div class="grid-cell project-date">${safePlanDate}</div>
                            <div class="grid-cell project-date">${safeEstimateDate}</div>
                            ${showDateIndicatorColumn ? `<div class="grid-cell project-indicator-cell">${dateIndicatorHtml}</div>` : ''}
                        `
                    }
                    <div class="grid-cell project-indicator-cell">
                        <div class="project-planfact-container">
                            <span class="project-planfact-value">${safeHoursDisplay}</span>
                        </div>
                    </div>
                    <div class="grid-cell project-indicator-cell">
                        <span class="project-value-badge project-value-badge--${escapeHtml(hoursBadgeTone)}${hoursBadgeCaption ? ' project-value-badge--stacked' : ''}"${hoursBadgeTooltip}>
                            <span class="project-value-badge__value">${safeHoursBadgeValueDisplay}</span>
                            ${hoursBadgeCaption ? `<span class="project-value-badge__caption">${safeHoursBadgeCaption}</span>` : ''}
                        </span>
                    </div>
                </div>
            `;
        });
        
        html += `
                </div>
            </div>
        `;
    });
    
    html += `</div>`;
    return html;
}

function getProjectTableFactEndDelayDays(project) {
    if (
        !project
        || !(project.factEnd instanceof Date)
        || isNaN(project.factEnd.getTime())
        || !(project.planDate instanceof Date)
        || isNaN(project.planDate.getTime())
        || project.factEnd <= project.planDate
    ) {
        return 0;
    }

    return Math.max(1, Math.ceil((project.factEnd.getTime() - project.planDate.getTime()) / 86400000));
}

function getProjectDetailManagementGroups(project) {
    const managements = Array.isArray(project?.managements)
        ? project.managements.filter(Boolean)
        : [];

    if (managements.length > 0) {
        return managements;
    }

    return [project?.management || 'Без управления'];
}

function getProjectDetailCenterGroups(project) {
    const centers = Array.isArray(project?.centers)
        ? project.centers.map(normalizeProjectCenterLabel).filter(Boolean)
        : [];

    if (centers.length > 0) {
        return Array.from(new Set(centers)).sort(compareProjectTableCenters);
    }

    const center = normalizeProjectCenterLabel(project?.center);
    return center ? [center] : ['Без центра'];
}

// ============================================================
// НАВИГАЦИЯ
// ============================================================

function activateNavButton(sheetName) {
    document.querySelectorAll('.nav-btn').forEach(btn => {
        if (btn.dataset.sheet === sheetName) {
            btn.classList.add('active');
        } else {
            btn.classList.remove('active');
        }
    });

    const activeBtn = Array.from(document.querySelectorAll('.nav-btn'))
        .find(btn => btn.dataset.sheet === sheetName);
    const currentLabel = document.getElementById('nav-menu-current');
    if (activeBtn && currentLabel) {
        currentLabel.textContent = activeBtn.textContent.trim();
    }
}

function renderProductionProgramPlanFact() {
    const container = document.getElementById('production-program-projects-container');
    if (!container) return;

    const projectsTitle = document.querySelector('#production-program-projects-block > h3');
    if (projectsTitle) projectsTitle.textContent = 'План-факт по проектам';

    try {
        const cacheKey = getProductionProgramRenderCacheKey();
        if (workbookDataCache.productionProgramHtml.has(cacheKey)) {
            container.innerHTML = workbookDataCache.productionProgramHtml.get(cacheKey);
            return;
        }

        const projects = getTechBlockProjectPlanFactData();
        const html = renderTechBlockProjectPlanFact(projects);
        workbookDataCache.productionProgramHtml.set(cacheKey, html);
        container.innerHTML = html;
    } catch (error) {
        console.error('Ошибка отрисовки блока План-факт по проектам:', error);
        container.innerHTML = `
            <div class="planfact-projects planfact-projects--empty">
                <div class="planfact-empty-state">Не удалось загрузить план-факт по проектам. Обновите данные или проверьте структуру таблиц техблока.</div>
            </div>
        `;
    }
}

function getProductionProgramRenderCacheKey() {
    return JSON.stringify({
        filters: productionProgramFilters,
        fact: productionProgramFactVisible,
        details: productionProgramDetailsExpanded
    });
}

function renderCalculationsAudit() {
    const container = document.getElementById('calculations-container');
    if (!container) return;

    try {
        const data = getCalculationsAuditData();
        container.innerHTML = `
            <div class="calc-audit">
                <div class="calc-audit__intro">
                    <div>
                        <h2>Самопроверка расчетов по плоским таблицам</h2>
                        <p>Здесь показано, какие листы используются, по каким правилам считаются ключевые показатели и какие суммы получаются по центрам за контрольные месяцы.</p>
                    </div>
                    <div class="calc-audit__date">
                        <span>контрольная дата</span>
                        <strong>${escapeHtml(formatDateFull(data.referenceDate))}</strong>
                    </div>
                </div>
                ${renderCalculationSourceCards(data.sources)}
                ${renderCalculationFormulaCards()}
                ${renderCalculationCenterSummary(data)}
                ${renderCalculationProjectBreakdown(data.projects)}
            </div>
        `;
    } catch (error) {
        console.error('Ошибка отрисовки вкладки расчетов:', error);
        container.innerHTML = `
            <div class="calc-audit calc-audit--empty">
                <div class="planfact-empty-state">Не удалось собрать расчеты. Проверьте наличие плоских таблиц в Excel.</div>
            </div>
        `;
    }
}

function getCalculationsAuditData() {
    const sourceSheetNames = ['Проекты_ТехБлок', 'Загрузка_ТехБлок', 'Сотрудники_ТехБлок', 'Табель', getSalesPlanWorksheetName() || 'ПП', 'Справочник'];
    const sources = sourceSheetNames.map(getCalculationSourceInfo);
    const projects = getTechBlockProjectPlanFactDataFromTables();
    const referenceDate = getReferenceDate();
    const currentMonthKey = getMonthKey(new Date(referenceDate.getFullYear(), referenceDate.getMonth(), 1));
    const nextMonthDate = new Date(referenceDate.getFullYear(), referenceDate.getMonth() + 1, 1);
    const nextMonthKey = getMonthKey(nextMonthDate);
    const monthIntervals = [
        { start: new Date(referenceDate.getFullYear(), referenceDate.getMonth(), 1) },
        { start: nextMonthDate }
    ];
    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);
    const centerOrder = ['ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ', 'ЛОИ', 'УВП'];
    const centerRows = centerOrder.map(center => getCalculationCenterRow(projects, center, currentMonthKey, nextMonthKey, capacityByCenterMonth));

    return {
        referenceDate,
        currentMonthKey,
        currentMonthLabel: formatMonthName(referenceDate.getMonth()),
        nextMonthKey,
        nextMonthLabel: formatMonthName(nextMonthDate.getMonth()),
        sources,
        centers: centerRows.filter(row => row.hasData),
        projects: getCalculationProjectRows(projects, currentMonthKey, nextMonthKey)
    };
}

function getCalculationSourceInfo(sheetName) {
    const sheet = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
    if (!sheet || !sheet['!ref'] || typeof XLSX === 'undefined') {
        return { sheetName, rows: 0, columns: 0, headers: [], status: 'нет листа' };
    }

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const headers = [];

    for (let col = range.s.c; col <= range.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const value = sheet[address] ? sheet[address].v : '';
        if (isValid(value)) headers.push(String(value).trim());
    }

    return {
        sheetName,
        rows: Math.max(0, range.e.r - range.s.r),
        columns: Math.max(0, range.e.c - range.s.c + 1),
        headers,
        status: headers.length > 0 ? 'готово' : 'нет заголовков'
    };
}

function getCalculationCenterRow(projects, center, currentMonthKey, nextMonthKey, capacityByCenterMonth) {
    const capacity = capacityByCenterMonth.get(center)?.get(currentMonthKey) || null;
    const nextCapacity = capacityByCenterMonth.get(center)?.get(nextMonthKey) || null;
    let activeProjectCount = 0;
    let plannedProjectCount = 0;
    let outsideProjectCount = 0;
    let currentFactHours = 0;
    let nextForecastHours = 0;
    const seenActiveProjects = new Set();

    projects.forEach(project => {
        const centers = getProjectCentersForLoad(project);
        if (!centers.includes(center)) return;

        if (project.isPlanned) {
            plannedProjectCount += 1;
        } else {
            outsideProjectCount += 1;
        }

        if (project.status.type === 'active') {
            seenActiveProjects.add(project.id);
            const monthMap = project.actualHoursByCenterMonth instanceof Map
                ? project.actualHoursByCenterMonth.get(center)
                : null;
            if (monthMap instanceof Map) {
                currentFactHours += monthMap.get(currentMonthKey) || 0;
            }
        }

        if (project.status.type === 'done' || project.status.type === 'cancelled') return;

        const planMonthMap = project.planHoursByCenterMonth instanceof Map
            ? project.planHoursByCenterMonth.get(center)
            : null;

        if (planMonthMap instanceof Map) {
            nextForecastHours += planMonthMap.get(nextMonthKey) || 0;
        }
    });

    activeProjectCount = seenActiveProjects.size;

    const currentAvailableHours = capacity ? capacity.availableHours || 0 : 0;
    const nextAvailableHours = nextCapacity ? nextCapacity.availableHours || 0 : 0;

    return {
        center,
        centerName: getProjectStageFullLabel(center),
        activeProjectCount,
        plannedProjectCount,
        outsideProjectCount,
        currentFactHours,
        currentAvailableHours,
        currentLoadPercent: currentAvailableHours > 0 ? Math.round((currentFactHours / currentAvailableHours) * 100) : null,
        nextForecastHours,
        nextAvailableHours,
        nextForecastPercent: nextAvailableHours > 0 ? Math.round((nextForecastHours / nextAvailableHours) * 100) : null,
        hasData: activeProjectCount > 0 || plannedProjectCount > 0 || outsideProjectCount > 0 || currentFactHours > 0 || nextForecastHours > 0 || currentAvailableHours > 0
    };
}

function getCalculationProjectRows(projects, currentMonthKey, nextMonthKey) {
    return projects
        .filter(project => project.status.type !== 'done' && project.status.type !== 'cancelled')
        .slice(0, 80)
        .map(project => {
            let currentFactHours = 0;
            let nextForecastHours = 0;

            (project.actualHoursByCenterMonth || new Map()).forEach(monthMap => {
                if (monthMap instanceof Map) currentFactHours += monthMap.get(currentMonthKey) || 0;
            });

            (project.planHoursByCenterMonth || new Map()).forEach(monthMap => {
                if (monthMap instanceof Map) nextForecastHours += monthMap.get(nextMonthKey) || 0;
            });

            return {
                id: project.id,
                name: project.name,
                center: project.center || getProjectCentersForLoad(project).join(', '),
                status: project.status.label,
                planFlag: project.isPlanned ? 'из плана' : 'вне плана',
                planHours: project.planHours || 0,
                actualHours: project.actualHours || 0,
                currentFactHours,
                forecastHours: nextForecastHours,
                nextForecastHours,
                factSource: project.actualHours > 0 ? 'Загрузка_ТехБлок' : 'Проекты_ТехБлок',
                planSource: project.isPlanned ? 'ПП' : 'Проекты_ТехБлок'
            };
        });
}

function renderCalculationSourceCards(sources) {
    return `
        <section class="calc-section">
            <div class="calc-section__header">
                <h3>Плоские таблицы</h3>
                <span>считываются по заголовкам первой строки</span>
            </div>
            <div class="calc-source-grid">
                ${sources.map(source => `
                    <div class="calc-source-card">
                        <div class="calc-source-card__top">
                            <strong>${escapeHtml(source.sheetName)}</strong>
                            <span>${escapeHtml(source.status)}</span>
                        </div>
                        <div class="calc-source-card__meta">
                            <span>${formatHoursValue(source.rows)} строк</span>
                            <span>${formatHoursValue(source.columns)} столбцов</span>
                        </div>
                        <div class="calc-source-card__headers">${escapeHtml(source.headers.slice(0, 7).join(' · '))}${source.headers.length > 7 ? ' · ...' : ''}</div>
                    </div>
                `).join('')}
            </div>
        </section>
    `;
}

function renderCalculationFormulaCards() {
    const formulas = [
        {
            title: 'Факт по списанным часам',
            formula: 'сумма Загрузка_ТехБлок[Часы]',
            note: 'Для прошедших и текущего месяцев берутся списанные часы центра. Исключаются строки отсутствий и все строки с признаком Прогноз.'
        },
        {
            title: 'Фонд часов',
            formula: 'Табель[Часы] × Сотрудники_ТехБлок[Ставка] − часы отсутствий',
            note: 'Административный персонал исключается. Дата выхода и дата увольнения ограничивают период работы сотрудника.'
        },
        {
            title: 'Прогноз / очередь',
            formula: 'ПП[Часы] по месяцам',
            note: 'Показывается начиная со следующего месяца. Строки Загрузка_ТехБлок с признаком Прогноз в расчет не добавляются.'
        },
        {
            title: 'Плановые часы',
            formula: 'ПП[Часы] + ПП[Месяц] + ПП[Управление] или Проекты_ТехБлок[Количество ПЛАНовых часов]',
            note: 'Для проектов из плана продаж часы берутся из ПП, раскладываются по месяцу строки и центру/управлению. Для проектов вне плана — из Проекты_ТехБлок.'
        }
    ];

    return `
        <section class="calc-section">
            <div class="calc-section__header">
                <h3>Правила расчета</h3>
                <span>короткая версия формул</span>
            </div>
            <div class="calc-formula-grid">
                ${formulas.map(item => `
                    <div class="calc-formula-card">
                        <strong>${escapeHtml(item.title)}</strong>
                        <code>${escapeHtml(item.formula)}</code>
                        <p>${escapeHtml(item.note)}</p>
                    </div>
                `).join('')}
            </div>
        </section>
    `;
}

function renderCalculationCenterSummary(data) {
    return `
        <section class="calc-section">
            <div class="calc-section__header">
                <h3>Сверка по центрам</h3>
                <span>${escapeHtml(data.currentMonthLabel)}: факт в работе, ${escapeHtml(data.nextMonthLabel)}: прогноз / очередь</span>
            </div>
            <div class="calc-table-wrap">
                <div class="calc-table calc-table--centers">
                    <div class="calc-table__row calc-table__row--head">
                        <span>Центр</span>
                        <span>Активные проекты</span>
                        <span>Из плана</span>
                        <span>Вне плана</span>
                        <span>Факт, ч</span>
                        <span>Фонд, ч</span>
                        <span>Загрузка</span>
                        <span>Прогноз, ч</span>
                        <span>Прогноз</span>
                    </div>
                    ${data.centers.map(row => `
                        <div class="calc-table__row">
                            <span><strong>${escapeHtml(row.center)}</strong><small>${escapeHtml(row.centerName)}</small></span>
                            <span>${row.activeProjectCount}</span>
                            <span>${row.plannedProjectCount}</span>
                            <span>${row.outsideProjectCount}</span>
                            <span>${formatHoursValue(row.currentFactHours)}</span>
                            <span>${formatHoursValue(row.currentAvailableHours)}</span>
                            <span class="calc-percent">${row.currentLoadPercent === null ? '-' : `${row.currentLoadPercent}%`}</span>
                            <span>${formatHoursValue(row.nextForecastHours)}</span>
                            <span class="calc-percent calc-percent--forecast">${row.nextForecastPercent === null ? '-' : `${row.nextForecastPercent}%`}</span>
                        </div>
                    `).join('')}
                </div>
            </div>
        </section>
    `;
}

function renderCalculationProjectBreakdown(projects) {
    return `
        <section class="calc-section">
            <details class="calc-details">
                <summary>
                    <span>Расшифровка по проектам</span>
                    <small>${projects.length} строк для проверки</small>
                </summary>
                <div class="calc-table-wrap">
                    <div class="calc-table calc-table--projects">
                        <div class="calc-table__row calc-table__row--head">
                            <span>ID</span>
                            <span>Проект</span>
                            <span>Центр</span>
                            <span>Статус</span>
                            <span>Метка</span>
                            <span>План, ч</span>
                            <span>Факт, ч</span>
                            <span>Факт текущего месяца</span>
                            <span>Прогноз следующего месяца</span>
                            <span>Источник плана</span>
                            <span>Источник факта</span>
                        </div>
                        ${projects.map(project => `
                            <div class="calc-table__row">
                                <span>${escapeHtml(project.id)}</span>
                                <span>${escapeHtml(project.name)}</span>
                                <span>${escapeHtml(project.center || '-')}</span>
                                <span>${escapeHtml(project.status || '-')}</span>
                                <span>${escapeHtml(project.planFlag)}</span>
                                <span>${formatHoursValue(project.planHours)}</span>
                                <span>${formatHoursValue(project.actualHours)}</span>
                                <span>${formatHoursValue(project.currentFactHours)}</span>
                                <span>${formatHoursValue(project.nextForecastHours)}</span>
                                <span>${escapeHtml(project.planSource)}</span>
                                <span>${escapeHtml(project.factSource)}</span>
                            </div>
                        `).join('')}
                    </div>
                </div>
            </details>
        </section>
    `;
}

function getTechBlockProjectPlanFactData() {
    try {
        const tableProjects = getTechBlockProjectPlanFactDataFromTables();
        if (tableProjects.length > 0) return tableProjects;
    } catch (error) {
        console.error('Ошибка сборки план-факта из таблиц техблока:', error);
    }

    return getTechBlockProjectPlanFactDataFromLegacySheet();
}

function getTechBlockProjectPlanFactDataFromLegacySheet() {
    const projectSheet = workbook && workbook.Sheets ? workbook.Sheets['Проекты'] : null;
    if (!projectSheet) return [];

    const projectCache = createWorksheetCache(projectSheet);
    const projects = [];
    const maxRows = 500;
    const centerStageColumns = ['E', 'F', 'G', 'H'];
    const centerStageHeaders = centerStageColumns.map(column => ({
        column,
        shortLabel: isValid(getCellValue(`${column}1`, projectCache))
            ? String(getCellValue(`${column}1`, projectCache)).trim()
            : column
    }));

    for (let row = 2; row <= maxRows; row++) {
        const id = getCellValue(`A${row}`, projectCache);
        const name = getCellValue(`B${row}`, projectCache);

        if (!isValid(id) && !isValid(name)) {
            continue;
        }

        const centerStages = centerStageHeaders
            .map(stage => ({
                label: getProjectStageFullLabel(stage.shortLabel),
                shortLabel: stage.shortLabel,
                rawStatus: getCellValue(`${stage.column}${row}`, projectCache)
            }))
            .filter(stage => isFilledProjectStageValue(stage.rawStatus))
            .map(stage => ({
                label: stage.label,
                shortLabel: stage.shortLabel,
                status: normalizeProjectStage(stage.rawStatus)
            }));

        const factStartRaw = getCellValue(`J${row}`, projectCache);
        const factEndRaw = getCellValue(`K${row}`, projectCache);
        const actualHoursRaw = getCellValue(`L${row}`, projectCache);
        const planSalesRaw = getCellValue(`M${row}`, projectCache);
        const planStartRaw = getCellValue(`N${row}`, projectCache);
        const planEndRaw = getCellValue(`O${row}`, projectCache);
        const planHoursRaw = getCellValue(`P${row}`, projectCache);

        const factStart = parseDateValue(factStartRaw);
        const factEnd = parseDateValue(factEndRaw);
        const planStart = parseDateValue(planStartRaw);
        const planEnd = parseDateValue(planEndRaw);
        const actualHours = parseNumericValue(actualHoursRaw) || 0;
        const planHours = parseNumericValue(planHoursRaw) || 0;
        const planSalesText = isValid(planSalesRaw) ? String(planSalesRaw).trim() : '';
        const isPlanned = planSalesText.toLowerCase() === 'да' || Boolean(planStart || planEnd || planHours > 0);

        projects.push({
            id: isValid(id) ? String(id).trim() : '-',
            name: isValid(name) ? String(name).trim() : 'Без названия',
            center: isValid(getCellValue(`Q${row}`, projectCache)) ? String(getCellValue(`Q${row}`, projectCache)).trim() : '',
            centerStages,
            stages: [
                { label: getProjectStageFullLabel('ЛОИ'), shortLabel: 'ЛОИ', status: normalizeProjectStage(getCellValue(`C${row}`, projectCache)) },
                { label: getProjectStageFullLabel('УВП'), shortLabel: 'УВП', status: normalizeProjectStage(getCellValue(`D${row}`, projectCache)) },
                ...centerStages
            ],
            status: normalizeProjectStage(getCellValue(`I${row}`, projectCache)),
            factStart,
            factEnd,
            factStartText: formatExcelDate(factStartRaw) || '-',
            factEndText: formatExcelDate(factEndRaw) || '-',
            actualHours,
            planSales: planSalesText,
            isPlanned,
            planStart,
            planEnd,
            planStartText: formatExcelDate(planStartRaw) || '-',
            planEndText: formatExcelDate(planEndRaw) || '-',
            planHours
        });
    }

    return projects
        .filter(isProjectVisibleInPlanFactYear)
        .map(project => ({
            ...project,
            planFlag: getProjectPlanFlag(project),
            deviation: getProjectPlanFactDeviation(project)
        }))
        .sort((a, b) => {
            const priority = { risk: 0, unstarted: 1, active: 2, ontime: 3, done: 4, neutral: 5 };
            const aPriority = priority[a.deviation.type] ?? 9;
            const bPriority = priority[b.deviation.type] ?? 9;
            if (aPriority !== bPriority) return aPriority - bPriority;

            const aDate = a.planEnd || a.factEnd || a.factStart;
            const bDate = b.planEnd || b.factEnd || b.factStart;
            if (aDate && bDate) return aDate - bDate;
            if (aDate) return -1;
            if (bDate) return 1;
            return a.name.localeCompare(b.name, 'ru');
        });
}

function getTechBlockProjectPlanFactDataFromTables() {
    if (workbookDataCache.projectPlanFact) {
        return workbookDataCache.projectPlanFact;
    }

    if (!workbook || !workbook.Sheets || !workbook.Sheets['Проекты_ТехБлок']) {
        return [];
    }

    const projectRows = getWorksheetRowsByHeaders('Проекты_ТехБлок');
    if (projectRows.length === 0) return [];

    const employeeLookup = getTechBlockEmployeeLookup();
    const loadByProject = getTechBlockLoadByProject(employeeLookup);
    const timesheetCalendar = getTechBlockTimesheetCalendar();
    const managementCenterLookup = getTechBlockManagementCenterLookup(projectRows);
    const salesPlanByProject = getTechBlockSalesPlanByProject(managementCenterLookup);

    const projectGroups = new Map();

    projectRows.forEach((row, index) => {
        const rawId = getRowValue(row, ['ID']);
        const id = normalizeProjectId(rawId);
        const name = getCleanTextValue(getRowValue(row, ['Наименование проекта', 'Проект']));
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const management = normalizeManagementLabel(getRowValue(row, ['Управление']));
        const key = id || `row-${index}-${center || management || 'project'}-${name || 'empty'}`;

        if (!projectGroups.has(key)) {
            projectGroups.set(key, {
                id,
                sourceRows: [],
                stageMap: new Map(),
                planHoursByCenter: new Map(),
                planHoursByWorkUnit: new Map(),
                actualHoursByCenterFromProjects: new Map(),
                actualHoursByWorkUnitFromProjects: new Map()
            });
        }

        const group = projectGroups.get(key);
        group.sourceRows.push(row);

        const stageCenter = center;
        if (stageCenter) {
            const status = normalizeProjectStage(getRowValue(row, ['Статус']));
            const existingStage = group.stageMap.get(stageCenter);
            const planHours = getRowNumber(row, ['Количество ПЛАНовых часов', 'Плановые часы']);
            const actualHours = getRowNumber(row, ['Количество ФАКТических часов', 'Фактические часы']);
            const workUnitKey = getProjectWorkUnitKey(center, management);
            const nextStage = {
                label: getProjectStageFullLabel(stageCenter),
                shortLabel: stageCenter,
                status,
                priority: getProjectStatusPriority(status.rawLabel || status.label)
            };

            if (!existingStage || nextStage.priority < existingStage.priority) {
                group.stageMap.set(stageCenter, nextStage);
            }

            addToNumericMap(group.planHoursByCenter, stageCenter, planHours);
            addToNumericMap(group.actualHoursByCenterFromProjects, stageCenter, actualHours);
            addToNumericMap(group.planHoursByWorkUnit, workUnitKey, planHours);
            addToNumericMap(group.actualHoursByWorkUnitFromProjects, workUnitKey, actualHours);
        }
    });

    salesPlanByProject.forEach((salesPlan, id) => {
        if (projectGroups.has(id)) return;

        projectGroups.set(id, {
            id,
            sourceRows: [],
            stageMap: new Map(),
            planHoursByCenter: new Map(),
            planHoursByWorkUnit: new Map(),
            actualHoursByCenterFromProjects: new Map(),
            actualHoursByWorkUnitFromProjects: new Map()
        });
    });

    const projects = Array.from(projectGroups.values()).map(group => {
        const rows = group.sourceRows;
        const firstRow = rows[0] || {};
        const id = group.id || '-';
        const salesPlan = group.id ? salesPlanByProject.get(group.id) : null;
        const load = group.id ? loadByProject.get(group.id) : null;
        const rowStatuses = rows.map(row => getRowValue(row, ['Статус']));
        const status = getOverallProjectStatus(rowStatuses);
        const stages = getOrderedProjectStages(Array.from(group.stageMap.values()));
        const centerStages = stages.filter(stage => stage.status && stage.status.type !== 'empty');
        const salesPlanCenters = salesPlan && salesPlan.planHoursByCenter instanceof Map
            ? Array.from(salesPlan.planHoursByCenter.keys())
            : [];
        const centers = Array.from(new Set([...centerStages.map(stage => stage.shortLabel), ...salesPlanCenters]));
        const fallbackPlanHours = sumMapValues(group.planHoursByCenter);
        const planHours = salesPlan ? salesPlan.planHours : fallbackPlanHours;
        const planHoursByCenter = salesPlan && salesPlan.planHoursByCenter?.size
            ? salesPlan.planHoursByCenter
            : getPlanHoursByCenterForProject(group.planHoursByCenter, centers, planHours, Boolean(salesPlan));
        const projectActualHours = sumMapValues(group.actualHoursByCenterFromProjects);
        const loadActualHours = load ? load.actualHours : 0;
        const actualHours = salesPlan ? loadActualHours : (loadActualHours > 0 ? loadActualHours : projectActualHours);
        const actualHoursByCenter = salesPlan
            ? (load?.centerHours || new Map())
            : load && load.centerHours.size > 0
                ? load.centerHours
                : group.actualHoursByCenterFromProjects;
        const projectDates = getProjectDatesByWorkUnits(rows, load);
        const planEnd = salesPlan?.planEnd || projectDates.contractEnd;
        const forecastEnd = projectDates.forecastEnd;
        const factStart = projectDates.factStart;
        const factEnd = projectDates.factEnd;
        const planStart = salesPlan?.planStart || getProjectPlanStartFromFacts(factStart, factEnd, planEnd, forecastEnd);
        const name = getCleanTextValue(getFirstValidRowValue(rows, ['Наименование проекта', 'Проект'])) || salesPlan?.name || (salesPlan ? 'Проект из плана продаж' : 'Без названия');
        const projectType = getCleanTextValue(getFirstValidRowValue(rows, ['Вид проекта'])) || '';
        const isPlanned = Boolean(salesPlan);
        const employeeCount = load ? load.employeeNames.size : getRowNumber(firstRow, ['Задействовано сотрудников']);
        const workingDayCount = countTimesheetDaysInRange(timesheetCalendar, factStart || planStart, factEnd || planEnd || forecastEnd);

        return {
            id,
            name,
            center: centers.join(', '),
            centerStages,
            stages,
            status,
            factStart,
            factEnd,
            factStartText: formatDateFull(factStart),
            factEndText: formatDateFull(factEnd),
            actualHours,
            actualHoursByCenter,
            actualHoursByCenterMonth: salesPlan
                ? (load?.centerMonthHours || new Map())
                : (load?.centerMonthHours || new Map()),
            forecastHoursByCenterMonth: load?.forecastCenterMonthHours || new Map(),
            employeeCount: employeeCount || 0,
            workingDayCount,
            projectType,
            planSales: salesPlan ? 'ПП' : (isPlanned ? 'да' : ''),
            salesPlan: salesPlan ? {
                businessDirections: Array.from(salesPlan.businessDirections),
                projectStatuses: Array.from(salesPlan.projectStatuses),
                costCenters: Array.from(salesPlan.costCenters),
                designBureaus: Array.from(salesPlan.designBureaus),
                managements: Array.from(salesPlan.managements)
            } : null,
            isPlanned,
            planStart,
            planEnd,
            forecastEnd,
            planStartText: formatDateFull(planStart),
            planEndText: formatDateFull(planEnd || forecastEnd),
            planHours,
            planHoursByCenter,
            planHoursByCenterMonth: salesPlan && salesPlan.planHoursByCenterMonth?.size
                ? salesPlan.planHoursByCenterMonth
                : salesPlan && salesPlan.planHoursByMonth?.size
                    ? getSalesPlanMonthHoursByCenterForProject(planHoursByCenter, salesPlan.planHoursByMonth)
                    : getPlanMonthHoursByCenterForProject(planHoursByCenter, planStart, planEnd || forecastEnd),
            forecastHours: load?.forecastHours || sumRowsByNumber(rows, ['Количество ПРОГНОЗируемых часов'])
        };
    });

    const preparedProjects = prepareTechBlockProjectPlanFactProjects(projects);
    workbookDataCache.projectPlanFact = preparedProjects;
    return preparedProjects;
}

function prepareTechBlockProjectPlanFactProjects(projects) {
    return projects
        .filter(isProjectIncludedInPlanFactCalculations)
        .filter(isProjectVisibleInPlanFactYear)
        .map(project => ({
            ...project,
            planFlag: getProjectPlanFlag(project),
            deviation: getProjectPlanFactDeviation(project)
        }))
        .sort((a, b) => {
            const priority = { risk: 0, unstarted: 1, active: 2, ontime: 3, done: 4, neutral: 5 };
            const aPriority = priority[a.deviation.type] ?? 9;
            const bPriority = priority[b.deviation.type] ?? 9;
            if (aPriority !== bPriority) return aPriority - bPriority;

            const aDate = a.planEnd || a.factEnd || a.factStart;
            const bDate = b.planEnd || b.factEnd || b.factStart;
            if (aDate && bDate) return aDate - bDate;
            if (aDate) return -1;
            if (bDate) return 1;
            return a.name.localeCompare(b.name, 'ru');
        });
}

function isProjectIncludedInPlanFactCalculations(project) {
    const type = project && project.status ? project.status.type : 'empty';
    return type !== 'done' && type !== 'cancelled';
}

function getWorksheetRowsByHeaders(sheetName) {
    if (workbookDataCache?.worksheetRows?.has(sheetName)) {
        return workbookDataCache.worksheetRows.get(sheetName);
    }

    const sheet = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
    if (!sheet || !sheet['!ref'] || typeof XLSX === 'undefined') return [];

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const cache = getCachedWorksheetCache(sheetName);
    const headers = [];

    for (let col = range.s.c; col <= range.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const header = getCleanTextValue(getCellValue(address, cache));
        headers.push(header || `Столбец ${col + 1}`);
    }

    const rows = [];

    for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex++) {
        const row = {};
        let hasData = false;

        headers.forEach((header, offset) => {
            const col = range.s.c + offset;
            const address = XLSX.utils.encode_cell({ r: rowIndex, c: col });
            const value = getCellValue(address, cache);
            row[header] = value;
            if (isFilledProjectStageValue(value)) hasData = true;
        });

        Object.defineProperties(row, {
            __sheetName: { value: sheetName },
            __rowIndex: { value: rowIndex }
        });

        if (hasData) rows.push(row);
    }

    workbookDataCache.worksheetRows.set(sheetName, rows);
    return rows;
}

function getRowValue(row, aliases) {
    const names = Array.isArray(aliases) ? aliases : [aliases];

    for (const name of names) {
        if (Object.prototype.hasOwnProperty.call(row, name) && isValid(row[name])) {
            return row[name];
        }
    }

    const normalizedNames = names.map(normalizeLookupText);
    const matchedKey = Object.keys(row).find(key => normalizedNames.includes(normalizeLookupText(key)));
    return matchedKey && isValid(row[matchedKey]) ? row[matchedKey] : '';
}

function getWorksheetRowCellValue(row, columnIndex) {
    if (!row || typeof row.__sheetName !== 'string' || typeof row.__rowIndex !== 'number') return '';
    if (typeof columnIndex !== 'number' || columnIndex < 0 || typeof XLSX === 'undefined') return '';

    const cache = getCachedWorksheetCache(row.__sheetName);
    const address = XLSX.utils.encode_cell({ r: row.__rowIndex, c: columnIndex });
    return getCellValue(address, cache);
}

function getFirstValidRowValue(rows, aliases) {
    for (const row of rows) {
        const value = getRowValue(row, aliases);
        if (isFilledProjectStageValue(value)) return value;
    }

    return '';
}

function getRowNumber(row, aliases) {
    const value = parseNumericValue(getRowValue(row, aliases));
    return value === null ? 0 : value;
}

function sumRowsByNumber(rows, aliases) {
    return rows.reduce((sum, row) => sum + getRowNumber(row, aliases), 0);
}

function getProjectNamesLookup() {
    if (workbookDataCache.projectNamesLookup instanceof Map) {
        return workbookDataCache.projectNamesLookup;
    }

    const lookup = new Map();
    const context = getWorksheetHeaderContext('Названия_проектов');
    if (!context) {
        workbookDataCache.projectNamesLookup = lookup;
        return lookup;
    }

    for (let rowIndex = context.range.s.r + 1; rowIndex <= context.range.e.r; rowIndex++) {
        const id = normalizeProjectIdForLookup(getCellValue(
            XLSX.utils.encode_cell({ r: rowIndex, c: 0 }),
            context.cache
        ));
        if (!id) continue;

        const name = getCleanTextValue(getWorksheetContextValue(context, rowIndex, [
            'Заказчик, название проекта',
            'Заказчик, название проектов'
        ]));
        if (!name) continue;

        if (!lookup.has(id)) {
            lookup.set(id, name);
        }
    }

    workbookDataCache.projectNamesLookup = lookup;
    return lookup;
}

function getCleanTextValue(value) {
    if (!isValid(value)) return '';
    const text = String(value).trim();
    return text === '-' || text === '—' ? '' : text;
}

function normalizeLookupText(value) {
    return getCleanTextValue(value).toLowerCase().replace(/\s+/g, ' ');
}

function normalizeProjectIdForLookup(value) {
    if (value === null || value === undefined) return '';

    const result = String(value).trim().replace(/\.0$/, '').replace(/\s+/g, '');
    return result === '0' || result === '-' || result === '—' ? '' : result;
}

function normalizeProjectId(value) {
    return normalizeProjectIdForLookup(value);
}

function normalizePersonName(value) {
    return normalizeLookupText(value).replace(/\./g, '').replace(/\s+/g, ' ');
}

function normalizeProjectCenterLabel(value) {
    const raw = getCleanTextValue(value);
    if (!raw) return '';

    const normalized = raw.toUpperCase().replace(/\s+/g, ' ').trim();
    const direct = {
        'ЦНГО': 'ЦНГО',
        'ЦДЦМ': 'ЦДЦМ',
        'ЦДПМ': 'ЦДПМ',
        'ЦОТ': 'ЦОТ',
        'УВП': 'УВП',
        'ЛОИ': 'ЛОИ'
    };

    if (direct[normalized]) return direct[normalized];

    const lower = raw.toLowerCase();
    if (lower.includes('нефтегазового оборудования')) return 'ЦНГО';
    if (lower.includes('центробежных насос') || lower.includes('центробежных компресс')) return 'ЦДЦМ';
    if (lower.includes('поршневых машин')) return 'ЦДПМ';
    if (lower.includes('осевых турбомашин')) return 'ЦОТ';
    if (lower.includes('виртуального полигона')) return 'УВП';
    if (lower.includes('лабораторн')) return 'ЛОИ';

    return raw;
}

function normalizeManagementLabel(value) {
    return getCleanTextValue(value);
}

function getProjectWorkUnitKey(center, management) {
    return `${normalizeProjectCenterLabel(center)}||${normalizeManagementLabel(management)}`;
}

function getProjectDatesByWorkUnits(rows, load) {
    let contractEnd = null;
    let forecastEnd = null;
    let factStart = null;
    let factEnd = null;

    rows.forEach(row => {
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const management = normalizeManagementLabel(getRowValue(row, ['Управление']));
        const workUnit = load?.workUnits.get(getProjectWorkUnitKey(center, management));
        const rowFactStart = parseDateValue(getRowValue(row, ['ФАКТ начало', 'Фактическая дата начала'])) || workUnit?.factStart || null;
        const rowFactEnd = parseDateValue(getRowValue(row, ['ФАКТ окончание', 'Фактическая дата окончания'])) || workUnit?.factEnd || null;

        contractEnd = getLaterDate(contractEnd, parseDateValue(getRowValue(row, [
            'План дата завершения этапа по графику проекта',
            'Дата завершения работ по договору',
            'Плановая дата окончания'
        ])));
        forecastEnd = getLaterDate(forecastEnd, parseDateValue(getRowValue(row, ['ПРОГНОЗ окончания', 'Прогнозируемая дата окончания'])));
        factStart = getEarlierDate(factStart, rowFactStart);
        factEnd = getLaterDate(factEnd, rowFactEnd || rowFactStart);
    });

    if (!factStart && load?.factStart) factStart = load.factStart;
    if (!factEnd && load?.factEnd) factEnd = load.factEnd;

    return { contractEnd, forecastEnd, factStart, factEnd };
}

function getOrderedProjectStages(stages) {
    const order = ['ЛОИ', 'УВП', 'ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ'];

    return stages
        .filter(stage => stage && stage.shortLabel)
        .sort((a, b) => {
            const aIndex = order.indexOf(a.shortLabel);
            const bIndex = order.indexOf(b.shortLabel);
            const aRank = aIndex === -1 ? order.length : aIndex;
            const bRank = bIndex === -1 ? order.length : bIndex;
            if (aRank !== bRank) return aRank - bRank;
            return a.shortLabel.localeCompare(b.shortLabel, 'ru');
        })
        .map(({ priority, ...stage }) => stage);
}

function getTechBlockEmployeeLookup() {
    if (workbookDataCache.employeeLookup) return workbookDataCache.employeeLookup;

    const rows = getWorksheetRowsByHeaders('Сотрудники_ТехБлок');
    const employees = new Map();

    rows.forEach(row => {
        const fio = normalizePersonName(getRowValue(row, ['ФИО']));
        if (!fio) return;

        employees.set(fio, {
            center: normalizeProjectCenterLabel(getRowValue(row, ['Центр'])),
            management: normalizeManagementLabel(getRowValue(row, ['Управление'])),
            terminationDate: parseDateValue(getRowValue(row, ['Дата увольнения'])),
            rate: getRowNumber(row, ['Ставка']) || 1,
            isAdmin: isAdminPersonnelValue(getRowValue(row, ['Административный персонал']))
        });
    });

    workbookDataCache.employeeLookup = employees;
    return employees;
}

function getWorksheetHeaderContext(sheetName) {
    const sheet = workbook && workbook.Sheets ? workbook.Sheets[sheetName] : null;
    if (!sheet || !sheet['!ref'] || typeof XLSX === 'undefined') return null;

    const range = XLSX.utils.decode_range(sheet['!ref']);
    const cache = getCachedWorksheetCache(sheetName);
    const headerMap = new Map();

    for (let col = range.s.c; col <= range.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: range.s.r, c: col });
        const header = normalizeLookupText(getCellValue(address, cache));
        if (header && !headerMap.has(header)) {
            headerMap.set(header, col);
        }
    }

    return { range, cache, headerMap };
}

function getWorksheetContextValue(context, rowIndex, aliases) {
    if (!context) return '';
    const names = Array.isArray(aliases) ? aliases : [aliases];

    for (const name of names) {
        const col = context.headerMap.get(normalizeLookupText(name));
        if (col === undefined) continue;

        const address = XLSX.utils.encode_cell({ r: rowIndex, c: col });
        const value = getCellValue(address, context.cache);
        if (isValid(value)) return value;
    }

    return '';
}

function isAdminPersonnelValue(value) {
    if (!isValid(value)) return false;
    const normalized = normalizeLookupText(value);
    return normalized === '1' || normalized === 'да' || normalized === 'true';
}

function isForecastLoadValue(value) {
    if (!isValid(value)) return false;
    const numeric = parseNumericValue(value);
    if (numeric !== null) return numeric !== 0;
    const normalized = normalizeLookupText(value);
    return normalized !== '' && normalized !== '0';
}

function isProjectAbsenceValue(value) {
    const normalized = normalizeLookupText(value);
    return normalized !== '' && normalized !== 'к';
}

function isEmployeeActiveOnDate(date, terminationDate) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return true;
    if (!(terminationDate instanceof Date) || isNaN(terminationDate.getTime())) return true;
    return date <= terminationDate;
}

function getTimesheetHoursForLoadRow(context, rowIndex, fallbackHours) {
    const timesheetHours = parseNumericValue(getWorksheetContextValue(context, rowIndex, ['Табель']));
    if (timesheetHours !== null) return timesheetHours;
    return fallbackHours || 0;
}

function getTechBlockLoadByProject(employeeLookup) {
    if (workbookDataCache.loadByProject) return workbookDataCache.loadByProject;

    const context = getWorksheetHeaderContext('Загрузка_ТехБлок');
    const loadMap = new Map();

    if (!context) {
        workbookDataCache.loadByProject = loadMap;
        return loadMap;
    }

    for (let rowIndex = context.range.s.r + 1; rowIndex <= context.range.e.r; rowIndex++) {
        const id = normalizeProjectId(getWorksheetContextValue(context, rowIndex, ['ID']));
        if (!id) continue;

        if (!loadMap.has(id)) {
            loadMap.set(id, {
                actualHours: 0,
                forecastHours: 0,
                absenceHours: 0,
                centerHours: new Map(),
                centerMonthHours: new Map(),
                forecastCenterMonthHours: new Map(),
                workUnitHours: new Map(),
                employeeNames: new Set(),
                workUnits: new Map(),
                factStart: null,
                factEnd: null
            });
        }

        const entry = loadMap.get(id);
        const hours = parseNumericValue(getWorksheetContextValue(context, rowIndex, ['Часы'])) || 0;
        const date = parseDateValue(getWorksheetContextValue(context, rowIndex, ['Дата']));
        const fio = getCleanTextValue(getWorksheetContextValue(context, rowIndex, ['ФИО']));
        const employee = employeeLookup.get(normalizePersonName(fio));
        const rowTerminationDate = parseDateValue(getWorksheetContextValue(context, rowIndex, ['Дата увольнения']));
        const terminationDate = rowTerminationDate || employee?.terminationDate || null;
        const isAdmin = isAdminPersonnelValue(getWorksheetContextValue(context, rowIndex, ['Административный персонал'])) || Boolean(employee?.isAdmin);
        const absence = getCleanTextValue(getWorksheetContextValue(context, rowIndex, ['Отсутствия']));
        const isAbsence = isProjectAbsenceValue(absence);
        const isForecast = isForecastLoadValue(getWorksheetContextValue(context, rowIndex, ['Прогноз']));
        const isActiveOnDate = isEmployeeActiveOnDate(date, terminationDate);
        const shouldUseForProjectHours = !isAdmin && isActiveOnDate && !isAbsence;
        const center = normalizeProjectCenterLabel(getWorksheetContextValue(context, rowIndex, ['Центр']))
            || employee?.center
            || '';
        const management = normalizeManagementLabel(getWorksheetContextValue(context, rowIndex, ['Управление']))
            || employee?.management
            || '';
        const workUnitKey = getProjectWorkUnitKey(center, management);

        if (shouldUseForProjectHours && fio) entry.employeeNames.add(fio);

        if (isAbsence && !isAdmin && isActiveOnDate) {
            entry.absenceHours += getTimesheetHoursForLoadRow(context, rowIndex, hours);
        }

        if (!shouldUseForProjectHours) continue;

        if (isForecast) {
            entry.forecastHours += hours;
            if (center && date) addNestedMapValue(entry.forecastCenterMonthHours, center, getMonthKey(date), hours);
            continue;
        }

        entry.actualHours += hours;
        if (center) addToNumericMap(entry.centerHours, center, hours);
        if (center && date) addNestedMapValue(entry.centerMonthHours, center, getMonthKey(date), hours);
        addToNumericMap(entry.workUnitHours, workUnitKey, hours);

        if (!entry.workUnits.has(workUnitKey)) {
            entry.workUnits.set(workUnitKey, {
                center,
                management,
                hours: 0,
                factStart: null,
                factEnd: null
            });
        }

        const unit = entry.workUnits.get(workUnitKey);
        unit.hours += hours;

        if (hours > 0 && date) {
            entry.factStart = getEarlierDate(entry.factStart, date);
            entry.factEnd = getLaterDate(entry.factEnd, date);
            unit.factStart = getEarlierDate(unit.factStart, date);
            unit.factEnd = getLaterDate(unit.factEnd, date);
        }
    }

    workbookDataCache.loadByProject = loadMap;
    return loadMap;
}

function getTechBlockTimesheetCalendar() {
    if (workbookDataCache.timesheetCalendar) return workbookDataCache.timesheetCalendar;

    const rows = getWorksheetRowsByHeaders('Табель');
    const calendar = new Map();

    rows.forEach(row => {
        const date = parseDateValue(getRowValue(row, ['Дата']));
        if (!date) return;

        calendar.set(date.toISOString().slice(0, 10), {
            date,
            weekday: getCleanTextValue(getRowValue(row, ['День недели'])),
            hours: getRowNumber(row, ['Часы'])
        });
    });

    workbookDataCache.timesheetCalendar = calendar;
    return calendar;
}

function getMonthIntervalCacheKey(monthIntervals) {
    if (!Array.isArray(monthIntervals) || monthIntervals.length === 0) return 'empty';

    return monthIntervals
        .map(interval => `${getMonthKey(interval.start)}:${getMonthKey(interval.end)}`)
        .join('|');
}

function getTechBlockCenterCapacityByMonth(monthIntervals) {
    const capacityCacheKey = getMonthIntervalCacheKey(monthIntervals);
    if (workbookDataCache.centerCapacityByRange.has(capacityCacheKey)) {
        return workbookDataCache.centerCapacityByRange.get(capacityCacheKey);
    }

    const capacityMap = new Map();
    const timesheetRows = getWorksheetRowsByHeaders('Табель');
    const timesheetDays = timesheetRows
        .map(row => ({
            date: parseDateValue(getRowValue(row, ['Дата'])),
            hours: getRowNumber(row, ['Часы'])
        }))
        .filter(day => day.date && day.hours > 0);
    const employeeRows = getWorksheetRowsByHeaders('Сотрудники_ТехБлок');
    const employeeByName = new Map();

    employeeRows.forEach(row => {
        const fio = normalizePersonName(getRowValue(row, ['ФИО']));
        if (!fio) return;

        employeeByName.set(fio, {
            center: normalizeProjectCenterLabel(getRowValue(row, ['Центр'])),
            startDate: parseDateValue(getRowValue(row, ['Дата выхода'])),
            terminationDate: parseDateValue(getRowValue(row, ['Дата увольнения'])),
            rate: getRowNumber(row, ['Ставка']) || 1,
            isAdmin: isAdminPersonnelValue(getRowValue(row, ['Административный персонал']))
        });
    });

    employeeByName.forEach(employee => {
        if (!employee.center || employee.isAdmin) return;

        timesheetDays.forEach(day => {
            if (!isEmployeeActiveInPeriod(day.date, employee.startDate, employee.terminationDate)) return;
            addCenterCapacityValue(capacityMap, employee.center, getMonthKey(day.date), 'grossHours', day.hours * employee.rate);
        });
    });

    const loadContext = getWorksheetHeaderContext('Загрузка_ТехБлок');
    if (loadContext) {
        for (let rowIndex = loadContext.range.s.r + 1; rowIndex <= loadContext.range.e.r; rowIndex++) {
            const absence = getWorksheetContextValue(loadContext, rowIndex, ['Отсутствия']);
            if (!isProjectAbsenceValue(absence)) continue;

            const fio = normalizePersonName(getWorksheetContextValue(loadContext, rowIndex, ['ФИО']));
            const employee = employeeByName.get(fio);
            const isAdmin = isAdminPersonnelValue(getWorksheetContextValue(loadContext, rowIndex, ['Административный персонал'])) || Boolean(employee?.isAdmin);
            if (isAdmin) continue;

            const date = parseDateValue(getWorksheetContextValue(loadContext, rowIndex, ['Дата']));
            const terminationDate = parseDateValue(getWorksheetContextValue(loadContext, rowIndex, ['Дата увольнения'])) || employee?.terminationDate || null;
            if (!isEmployeeActiveOnDate(date, terminationDate)) continue;

            const center = normalizeProjectCenterLabel(getWorksheetContextValue(loadContext, rowIndex, ['Центр'])) || employee?.center || '';
            if (!center || !date) continue;

            const rate = employee?.rate || 1;
            const absenceHours = getTimesheetHoursForLoadRow(loadContext, rowIndex, 0) * rate;
            addCenterCapacityValue(capacityMap, center, getMonthKey(date), 'absenceHours', absenceHours);
        }
    }

    capacityMap.forEach(monthMap => {
        monthMap.forEach(value => {
            value.availableHours = Math.max(0, (value.grossHours || 0) - (value.absenceHours || 0));
        });
    });

    workbookDataCache.centerCapacityByRange.set(capacityCacheKey, capacityMap);
    return capacityMap;
}

function addCenterCapacityValue(capacityMap, center, monthKey, field, value) {
    if (!center || !monthKey || !Number.isFinite(value) || value === 0) return;
    if (!capacityMap.has(center)) capacityMap.set(center, new Map());
    const monthMap = capacityMap.get(center);
    if (!monthMap.has(monthKey)) {
        monthMap.set(monthKey, { grossHours: 0, absenceHours: 0, availableHours: 0 });
    }
    const entry = monthMap.get(monthKey);
    entry[field] = (entry[field] || 0) + value;
}

function isEmployeeActiveInPeriod(date, startDate, terminationDate) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return false;
    if (startDate instanceof Date && !isNaN(startDate.getTime()) && date < startDate) return false;
    if (terminationDate instanceof Date && !isNaN(terminationDate.getTime()) && date > terminationDate) return false;
    return true;
}

function getTechBlockManagementCenterLookup(projectRows) {
    const lookup = new Map();
    const knownCenters = ['ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ', 'УВП', 'ЛОИ'];

    knownCenters.forEach(center => {
        lookup.set(normalizeLookupText(center), center);
        lookup.set(normalizeLookupText(getProjectStageFullLabel(center)), center);
    });

    projectRows.forEach(row => {
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const management = normalizeManagementLabel(getRowValue(row, ['Управление']));
        if (!center) return;

        lookup.set(normalizeLookupText(center), center);
        lookup.set(normalizeLookupText(getProjectStageFullLabel(center)), center);
        if (management) lookup.set(normalizeLookupText(management), center);
    });

    getWorksheetRowsByHeaders('Сотрудники_ТехБлок').forEach(row => {
        const center = normalizeProjectCenterLabel(getRowValue(row, ['Центр']));
        const management = normalizeManagementLabel(getRowValue(row, ['Управление']));
        if (!center || !management) return;

        const key = normalizeLookupText(management);
        if (!lookup.has(key)) lookup.set(key, center);
    });

    return lookup;
}

function getTechBlockSalesPlanByProject(managementCenterLookup = new Map()) {
    const sheetName = getSalesPlanWorksheetName();
    const rows = sheetName ? getWorksheetRowsByHeaders(sheetName) : [];
    const planMap = new Map();
    const seenRows = new Set();

    rows.forEach(row => {
        const id = normalizeProjectId(getRowValue(row, ['Номер задачи', 'ID']));
        if (!id) return;

        const planStart = parseDateValue(getRowValue(row, ['Дата начала проекта', 'Дата начала']));
        const planEnd = parseDateValue(getRowValue(row, ['Дата окончания проекта', 'Дата окончания']));
        const planHours = getRowNumber(row, ['Часы', 'Итого за 2026 год', 'Итого за 2025 год', 'Плановые часы', 'Количество ПЛАНовых часов']);
        const planMonthKey = getSalesPlanMonthKey(row, planStart, planEnd);
        const planCenter = getSalesPlanCenterFromRow(row, managementCenterLookup);
        const businessDirection = getCleanTextValue(getRowValue(row, ['Бизнес-направление']));
        const projectStatus = getCleanTextValue(getRowValue(row, ['Статус проекта']));
        const projectName = getCleanTextValue(getRowValue(row, ['Наименование проекта', 'Проект']));
        const costCenter = getCleanTextValue(getRowValue(row, ['Место возникновения затрат (МВЗ)', 'Место возникновения затрат', 'МВЗ']));
        const designBureau = getCleanTextValue(getRowValue(row, ['Наименование конструкторского бюро', 'Конструкторское бюро']));
        const management = getCleanTextValue(getRowValue(row, ['Управление']));
        const rowKey = [
            id,
            planStart ? planStart.getTime() : '',
            planEnd ? planEnd.getTime() : '',
            planMonthKey,
            planCenter,
            businessDirection,
            projectStatus,
            costCenter,
            designBureau,
            management,
            Math.round(planHours * 1000) / 1000
        ].join('|');

        if (seenRows.has(rowKey)) return;
        seenRows.add(rowKey);

        if (!planMap.has(id)) {
            planMap.set(id, {
                id,
                name: '',
                planStart: null,
                planEnd: null,
                planHours: 0,
                planHoursByCenter: new Map(),
                planHoursByMonth: new Map(),
                planHoursByCenterMonth: new Map(),
                businessDirections: new Set(),
                projectStatuses: new Set(),
                costCenters: new Set(),
                designBureaus: new Set(),
                managements: new Set(),
                rows: 0
            });
        }

        const entry = planMap.get(id);
        if (!entry.name && projectName) entry.name = projectName;
        entry.planStart = getEarlierDate(entry.planStart, planStart);
        entry.planEnd = getLaterDate(entry.planEnd, planEnd || planStart);
        entry.planHours += planHours;
        addToNumericMap(entry.planHoursByCenter, planCenter, planHours);
        addMonthMapValue(entry.planHoursByMonth, planMonthKey, planHours);
        addNestedMapValue(entry.planHoursByCenterMonth, planCenter, planMonthKey, planHours);
        addCleanValueToSet(entry.businessDirections, businessDirection);
        addCleanValueToSet(entry.projectStatuses, projectStatus);
        addCleanValueToSet(entry.costCenters, costCenter);
        addCleanValueToSet(entry.designBureaus, designBureau);
        addCleanValueToSet(entry.managements, management);
        entry.rows += 1;
    });

    return planMap;
}

function getSalesPlanWorksheetName() {
    if (!workbook || !workbook.Sheets) return '';
    return ['ПП', 'План_продаж', 'План-продаж'].find(sheetName => workbook.Sheets[sheetName]) || '';
}

function getSalesPlanCenterFromRow(row, managementCenterLookup) {
    const management = getCleanTextValue(getRowValue(row, ['Управление']));
    return resolveSalesPlanCenterValue(management, managementCenterLookup);
}

function resolveSalesPlanCenterValue(value, managementCenterLookup) {
    const normalized = normalizeLookupText(value);
    if (!normalized) return '';

    const direct = normalizeProjectCenterLabel(value);
    const directKey = normalizeLookupText(direct);
    if (managementCenterLookup instanceof Map && managementCenterLookup.has(directKey)) {
        return managementCenterLookup.get(directKey);
    }

    if (managementCenterLookup instanceof Map && managementCenterLookup.has(normalized)) {
        return managementCenterLookup.get(normalized);
    }

    return '';
}

function addCleanValueToSet(set, value) {
    if (!(set instanceof Set)) return;
    const cleanValue = getCleanTextValue(value);
    if (cleanValue) set.add(cleanValue);
}

function getSalesPlanMonthKey(row, planStart, planEnd) {
    const monthValue = getRowValue(row, ['Месяц', 'Месяц плана', 'Месяц загрузки']);
    const monthDate = parseSalesPlanMonthDate(monthValue, planStart, planEnd);
    return monthDate ? getMonthKey(monthDate) : '';
}

function parseSalesPlanMonthDate(value, planStart, planEnd) {
    if (!isValid(value)) return null;

    const fallbackYear = getSalesPlanMonthFallbackYear(planStart, planEnd);
    const numericMonth = parseNumericValue(value);
    const rawText = String(value).trim();
    if (
        numericMonth !== null
        && Number.isInteger(numericMonth)
        && numericMonth >= 1
        && numericMonth <= 12
        && !/[.\/-]/.test(rawText)
    ) {
        return new Date(fallbackYear, numericMonth - 1, 1);
    }

    const parsedDate = parseDateValue(value);
    if (parsedDate) {
        return new Date(parsedDate.getFullYear(), parsedDate.getMonth(), 1);
    }

    const monthIndex = parseRuMonthName(value);
    if (monthIndex === null) return null;

    const yearMatch = rawText.match(/\b(19\d{2}|20\d{2})\b/);
    const year = yearMatch ? parseInt(yearMatch[1], 10) : fallbackYear;
    return new Date(year, monthIndex, 1);
}

function getSalesPlanMonthFallbackYear(planStart, planEnd) {
    const startYear = planStart instanceof Date && !isNaN(planStart.getTime()) ? planStart.getFullYear() : null;
    const endYear = planEnd instanceof Date && !isNaN(planEnd.getTime()) ? planEnd.getFullYear() : null;

    if (startYear && endYear && startYear === endYear) return startYear;
    if (startYear === 2026 || endYear === 2026) return 2026;
    return endYear || startYear || 2026;
}

function getPlanHoursByCenterForProject(sourceMap, centers, totalPlanHours, useSalesPlan) {
    if (!useSalesPlan) return sourceMap;

    const result = new Map();
    const centerKeys = centers.length > 0
        ? centers
        : sourceMap instanceof Map
            ? Array.from(sourceMap.keys())
            : [];

    centerKeys.forEach(center => {
        result.set(center, totalPlanHours || 0);
    });

    return result;
}

function getSalesPlanMonthHoursByCenterForProject(planHoursByCenter, planHoursByMonth) {
    const result = new Map();
    if (!(planHoursByCenter instanceof Map) || planHoursByCenter.size === 0) return result;
    if (!(planHoursByMonth instanceof Map) || planHoursByMonth.size === 0) return result;

    const centerEntries = Array.from(planHoursByCenter.entries())
        .filter(([center]) => Boolean(center));
    if (centerEntries.length === 0) return result;

    const totalCenterHours = centerEntries.reduce((sum, [, hours]) => (
        sum + (Number.isFinite(hours) && hours > 0 ? hours : 0)
    ), 0);

    planHoursByMonth.forEach((monthHours, monthKey) => {
        if (!monthKey || !Number.isFinite(monthHours) || monthHours === 0) return;

        if (totalCenterHours > 0) {
            centerEntries.forEach(([center, centerHours]) => {
                if (!Number.isFinite(centerHours) || centerHours <= 0) return;
                addNestedMapValue(result, center, monthKey, monthHours * centerHours / totalCenterHours);
            });
            return;
        }

        const equalShare = monthHours / centerEntries.length;
        centerEntries.forEach(([center]) => addNestedMapValue(result, center, monthKey, equalShare));
    });

    return result;
}

function getPlanMonthHoursByCenterForProject(planHoursByCenter, planStart, planEnd) {
    const result = new Map();
    if (!(planHoursByCenter instanceof Map) || planHoursByCenter.size === 0) return result;

    const rangeStart = planStart || planEnd;
    const rangeEnd = planEnd || planStart;
    if (!(rangeStart instanceof Date) || isNaN(rangeStart.getTime())) return result;

    const months = getMonthKeysBetween(rangeStart, rangeEnd);
    if (months.length === 0) return result;

    planHoursByCenter.forEach((hours, center) => {
        if (!center || !hours) return;
        const monthlyHours = hours / months.length;
        months.forEach(monthKey => addNestedMapValue(result, center, monthKey, monthlyHours));
    });

    return result;
}

function countTimesheetDaysInRange(calendar, start, end) {
    const rangeStart = start || end;
    const rangeEnd = end || start;

    if (!(rangeStart instanceof Date) || isNaN(rangeStart.getTime())) return 0;

    return Array.from(calendar.values()).filter(day => {
        if (!day.date || day.hours <= 0) return false;
        return day.date >= rangeStart && day.date <= rangeEnd;
    }).length;
}

function addToNumericMap(map, key, value) {
    if (!key || !Number.isFinite(value) || value === 0) return;
    map.set(key, (map.get(key) || 0) + value);
}

function addNestedMapValue(map, outerKey, innerKey, value) {
    if (!outerKey || !innerKey || !Number.isFinite(value) || value === 0) return;
    if (!map.has(outerKey)) map.set(outerKey, new Map());
    const innerMap = map.get(outerKey);
    innerMap.set(innerKey, (innerMap.get(innerKey) || 0) + value);
}

function addMonthMapValue(map, monthKey, value) {
    if (!monthKey || !Number.isFinite(value) || value === 0) return;
    map.set(monthKey, (map.get(monthKey) || 0) + value);
}

function mergeNestedMonthMap(target, source, outerKey) {
    if (!(source instanceof Map) || !outerKey) return;
    const monthMap = source.get(outerKey);
    if (!(monthMap instanceof Map)) return;

    monthMap.forEach((value, monthKey) => {
        addMonthMapValue(target, monthKey, value);
    });
}

function sumMapValues(map) {
    return Array.from(map ? map.values() : []).reduce((sum, value) => sum + (Number.isFinite(value) ? value : 0), 0);
}

function getProjectPlanStartFromFacts(factStart, factEnd, planEnd, forecastEnd) {
    if (factStart) return factStart;
    if (factEnd) return factEnd;
    if (planEnd) return planEnd;
    return forecastEnd || null;
}

function getTechBlockStatusRules() {
    if (techBlockStatusRulesCache) return techBlockStatusRulesCache;

    const sheet = workbook && workbook.Sheets ? workbook.Sheets['Справочник'] : null;
    const rules = new Map();

    if (sheet) {
        const cache = createWorksheetCache(sheet);

        for (let row = 3; row <= 14; row++) {
            const raw = getCleanTextValue(getCellValue(`T${row}`, cache));
            const adapted = getCleanTextValue(getCellValue(`U${row}`, cache)) || raw;
            const priorityValue = parseNumericValue(getCellValue(`V${row}`, cache));

            if (!raw) continue;

            rules.set(normalizeLookupText(raw), {
                label: adapted,
                type: getProjectStatusType(adapted),
                priority: priorityValue === null ? 99 : priorityValue,
                rawLabel: raw
            });
        }
    }

    techBlockStatusRulesCache = rules;
    return rules;
}

function getOverallProjectStatus(rawStatuses) {
    const statuses = rawStatuses
        .filter(isFilledProjectStageValue)
        .map(status => normalizeProjectStage(status));

    if (statuses.length === 0) return normalizeProjectStage('');

    return statuses.sort((a, b) => {
        const aPriority = getProjectStatusPriority(a.rawLabel || a.label);
        const bPriority = getProjectStatusPriority(b.rawLabel || b.label);
        if (aPriority !== bPriority) return aPriority - bPriority;
        return a.label.localeCompare(b.label, 'ru');
    })[0];
}

function getProjectStatusPriority(value) {
    const rules = getTechBlockStatusRules();
    const rule = rules.get(normalizeLookupText(value));
    return rule ? rule.priority : 99;
}

function getProjectStatusType(label) {
    const normalized = normalizeLookupText(label);

    if (normalized === 'завершена' || normalized === 'завершен' || normalized === 'завершено') return 'done';
    if (normalized === 'в работе') return 'active';
    if (normalized === 'отменена' || normalized === 'отменен' || normalized === 'отменено') return 'cancelled';
    if (normalized === 'на паузе' || normalized === 'в очереди' || normalized === 'потенциальная' || normalized === 'у заказчика') return 'paused';

    return normalized ? 'empty' : 'empty';
}

function getProjectStageFullLabel(shortLabel) {
    const stageNames = {
        'ЦНГО': 'Центр нефтегазового оборудования',
        'ЦДЦМ': 'Центр центробежных насосов, центробежных компрессоров и роторных машин',
        'ЦДПМ': 'Центр деталей поршневых машин',
        'ЦОТ': 'Центр осевых турбомашин',
        'УВП': 'Управление виртуального полигона',
        'ЛОИ': 'Управление лабораторными исследованиями'
    };

    return stageNames[shortLabel] || shortLabel;
}

function isFilledProjectStageValue(value) {
    if (!isValid(value)) return false;

    const text = String(value).trim();
    return text !== '' && text !== '-' && text !== '—';
}

function getPlanFactDisplayRange() {
    return {
        startYear: 2025,
        endYear: 2027,
        startDate: new Date(2025, 6, 1),
        endDate: new Date(2027, 6, 1)
    };
}

function isDateInYear(date, year) {
    return date instanceof Date && !isNaN(date.getTime()) && date.getFullYear() === year;
}

function doesDateRangeIntersectYear(start, end, year) {
    const yearStart = new Date(year, 0, 1);
    const yearEnd = new Date(year, 11, 31, 23, 59, 59, 999);
    const rangeStart = start || end;
    const rangeEnd = end || start;

    if (!(rangeStart instanceof Date) || isNaN(rangeStart.getTime())) return false;

    return rangeStart <= yearEnd && rangeEnd >= yearStart;
}

function doesDateRangeIntersectDisplayRange(start, end) {
    const { startDate, endDate } = getPlanFactDisplayRange();
    const rangeStart = start || end;
    const rangeEnd = end || start;

    if (!(rangeStart instanceof Date) || isNaN(rangeStart.getTime())) return false;

    return rangeStart <= endDate && rangeEnd >= startDate;
}

function isProjectVisibleInPlanFactYear(project) {
    return doesDateRangeIntersectDisplayRange(project.planStart, project.planEnd)
        || doesDateRangeIntersectDisplayRange(project.factStart, project.factEnd);
}

function normalizeProjectStage(value) {
    const raw = isValid(value) ? String(value).trim() : '';
    const normalized = raw.toLowerCase();
    const rule = getTechBlockStatusRules().get(normalizeLookupText(raw));

    if (rule) {
        return {
            label: rule.label,
            type: rule.type,
            rawLabel: raw
        };
    }

    if (normalized === 'завершена' || normalized === 'завершен' || normalized === 'завершено') {
        return { label: 'Завершена', type: 'done', rawLabel: raw };
    }

    if (normalized === 'в работе') {
        return { label: 'В работе', type: 'active', rawLabel: raw };
    }

    if (normalized === 'на паузе' || normalized === 'в очереди' || normalized === 'потенциальная') {
        return { label: raw, type: 'paused', rawLabel: raw };
    }

    if (normalized === 'отменена' || normalized === 'отменен' || normalized === 'отменено') {
        return { label: raw, type: 'cancelled', rawLabel: raw };
    }

    return { label: raw || '-', type: 'empty', rawLabel: raw };
}

function getProjectPlanFactDeviation(project) {
    const referenceDate = getReferenceDate() || new Date();
    const hasFact = Boolean(project.factStart || project.factEnd || project.actualHours > 0);

    if (project.isPlanned && !hasFact) {
        return { label: 'не запущен', type: 'unstarted' };
    }

    if (!project.isPlanned) {
        return { label: '-', type: 'neutral' };
    }

    if (project.isPlanned && project.planEnd && project.factEnd) {
        const diffDays = Math.round((project.factEnd - project.planEnd) / 86400000);
        if (diffDays <= 0) {
            return { label: 'в срок', type: 'ontime' };
        }
        return { label: `+${diffDays} дн.`, type: 'risk' };
    }

    if (project.isPlanned && project.planEnd && !project.factEnd && referenceDate > project.planEnd) {
        const diffDays = Math.round((referenceDate - project.planEnd) / 86400000);
        return { label: `риск +${diffDays} дн.`, type: 'risk' };
    }

    if (project.status.type === 'done') {
        return { label: 'завершен', type: 'done' };
    }

    if (hasFact) {
        return { label: project.status.label || 'в работе', type: 'active' };
    }

    return { label: '-', type: 'neutral' };
}

function getProjectPlanFlag(project) {
    return project.isPlanned
        ? { label: 'из плана', type: 'planned' }
        : { label: 'вне плана', type: 'outside' };
}

function renderTechBlockProjectPlanFact(projects) {
    if (!projects || projects.length === 0) {
        return `
            <div class="planfact-projects planfact-projects--empty">
                <div class="planfact-empty-state">В таблицах техблока пока нет данных для отображения</div>
            </div>
        `;
    }

    const filteredProjects = applyProductionProgramFilters(projects);
    const filtersHtml = renderProductionProgramFilters(projects, filteredProjects.length);

    if (filteredProjects.length === 0) {
        return `
            <div class="planfact-projects">
                ${filtersHtml}
                <div class="planfact-empty-state">По выбранным фильтрам нет проектов</div>
            </div>
        `;
    }

    const total = filteredProjects.length;
    const plannedCount = filteredProjects.filter(project => project.isPlanned).length;
    const factCount = filteredProjects.filter(project => project.factStart || project.factEnd || project.actualHours > 0).length;
    const onTimeCount = filteredProjects.filter(project => project.deviation.type === 'ontime').length;
    const riskCount = filteredProjects.filter(project => project.deviation.type === 'risk').length;
    const unstartedCount = filteredProjects.filter(project => project.deviation.type === 'unstarted').length;
    const { startDate, endDate } = getTechBlockProjectTimelineRange(filteredProjects);
    const centerLoad = getProjectCenterLoadData(filteredProjects);

    return `
        <div class="planfact-projects">
            ${filtersHtml}
            ${renderPlanFactHeatmapSummary(filteredProjects, startDate, endDate)}
            ${renderProjectCenterLoadBlock(centerLoad, startDate, endDate)}
            ${renderPlanFactProjectDetailsBlock(filteredProjects, startDate, endDate, {
                total,
                plannedCount,
                factCount,
                onTimeCount,
                riskCount,
                unstartedCount
            })}
        </div>
    `;
}

function applyProductionProgramFilters(projects) {
    return projects.filter(project => (
        matchesProductionProgramFilter(project, 'businessDirection')
        && matchesProductionProgramFilter(project, 'projectStatus')
        && matchesProductionProgramFilter(project, 'costCenter')
        && matchesProductionProgramFilter(project, 'designBureau')
    ));
}

function matchesProductionProgramFilter(project, key) {
    const selected = productionProgramFilters[key];
    if (!selected) return true;

    const values = getProductionProgramFilterValues(project, key);
    return values.some(value => value === selected);
}

function getProductionProgramFilterValues(project, key) {
    const salesPlan = project.salesPlan || {};
    const map = {
        businessDirection: salesPlan.businessDirections,
        projectStatus: salesPlan.projectStatuses,
        costCenter: salesPlan.costCenters,
        designBureau: salesPlan.designBureaus
    };

    return Array.isArray(map[key]) ? map[key].filter(Boolean) : [];
}

function renderProductionProgramFilters(projects, filteredCount) {
    const configs = [
        { key: 'businessDirection', label: 'Бизнес-направление' },
        { key: 'projectStatus', label: 'Статус проекта' },
        { key: 'costCenter', label: 'МВЗ' },
        { key: 'designBureau', label: 'Конструкторское бюро' }
    ];

    return `
        <div class="production-program-filters" data-production-program-filters>
            <div class="production-program-filters__controls">
                ${configs.map(config => renderProductionProgramFilterControl(config, projects)).join('')}
                <button type="button" class="production-program-filters__reset" data-pp-filter-reset>Сбросить</button>
            </div>
            <div class="production-program-filters__fact">
                ${renderProductionProgramFactToggle()}
            </div>
        </div>
    `;
}

function renderProductionProgramFactToggle() {
    return `
        <label class="fact-toggle production-program-fact-toggle" title="Показать или скрыть фактические данные">
            <input class="fact-toggle__input" type="checkbox" data-pp-fact-toggle${productionProgramFactVisible ? ' checked' : ''}>
            <span class="fact-toggle__box" aria-hidden="true"></span>
            <span class="fact-toggle__label">Факт</span>
        </label>
    `;
}

function renderProductionProgramFilterControl(config, projects) {
    const options = getProductionProgramFilterOptions(projects, config.key);
    const selected = productionProgramFilters[config.key] || '';

    return `
        <label class="production-program-filter">
            <span>${escapeHtml(config.label)}</span>
            <select class="production-program-filter__select" data-pp-filter="${escapeHtml(config.key)}">
                <option value="">Все</option>
                ${options.map(value => `
                    <option value="${escapeHtml(value)}"${value === selected ? ' selected' : ''}>${escapeHtml(value)}</option>
                `).join('')}
            </select>
        </label>
    `;
}

function getProductionProgramFilterOptions(projects, key) {
    const values = new Set();
    projects.forEach(project => {
        getProductionProgramFilterValues(project, key).forEach(value => values.add(value));
    });

    return Array.from(values).sort((a, b) => a.localeCompare(b, 'ru', { numeric: true }));
}

function handleProductionProgramFilterChange(select) {
    const key = select.dataset.ppFilter;
    if (!Object.prototype.hasOwnProperty.call(productionProgramFilters, key)) return;

    productionProgramFilters[key] = select.value;
    renderProductionProgramPlanFact();
}

function resetProductionProgramFilters() {
    productionProgramFilters = {
        businessDirection: '',
        projectStatus: '',
        costCenter: '',
        designBureau: ''
    };
    renderProductionProgramPlanFact();
}

function getProductionProgramFactHiddenClass() {
    return productionProgramFactVisible ? '' : ' fact-hidden';
}

function handleProductionProgramFactToggle(input) {
    productionProgramFactVisible = input.checked;

    const content = document.getElementById('production-program-content');
    if (!content) return;

    content.querySelectorAll('[data-fact-toggle-block]').forEach(block => {
        block.classList.toggle('fact-hidden', !productionProgramFactVisible);
    });
}

function renderPlanFactProjectDetailsBlock(projects, startDate, endDate, summary) {
    const detailsRowsHtml = productionProgramDetailsExpanded
        ? projects.map(project => renderPlanFactProjectRow(project, startDate, endDate)).join('')
        : '<div class="planfact-details__lazy-note">Откройте блок, чтобы показать список проектов</div>';

    return `
        <details class="planfact-details${getProductionProgramFactHiddenClass()}" data-fact-toggle-block="planfact-details"${productionProgramDetailsExpanded ? ' open' : ''}>
            <summary class="planfact-details__summary">
                <span class="planfact-details__title">Сверка по каждому проекту</span>
                <span class="planfact-details__meta">${summary.total} ${formatProjectWord(summary.total)}</span>
            </summary>
            <div class="planfact-details__content">
                <div class="planfact-summary">
                    ${renderPlanFactSummaryCard('План продаж', summary.plannedCount, 'проектов в плане', 'blue')}
                    ${renderPlanFactSummaryCard('Есть в факте', summary.factCount, `из ${summary.total} проектов`, 'green', 'data-fact-field')}
                    ${renderPlanFactSummaryCard('В срок', summary.onTimeCount, 'по плановой дате', 'green', 'data-fact-field')}
                    ${renderPlanFactSummaryCard('Риски', summary.riskCount, 'сдвиг сроков', 'amber', 'data-fact-field')}
                    ${renderPlanFactSummaryCard('Не запущены', summary.unstartedCount, 'только в плане', 'red', 'data-fact-field')}
                </div>
                <div class="planfact-header">
                    <div>
                        <div class="planfact-title">Сверка плана продаж <span data-fact-field>и фактического выполнения</span></div>
                        <div class="planfact-subtitle">Синяя линия — план продаж из листа «ПП»<span data-fact-field>, зелёная/жёлтая — факт из таблиц техблока</span>.</div>
                    </div>
                    <div class="planfact-header__actions">
                        <div class="planfact-legend" data-fact-field>
                            <span><i class="planfact-dot planfact-dot--done"></i>завершена</span>
                            <span><i class="planfact-dot planfact-dot--active"></i>в работе</span>
                            <span><i class="planfact-dot planfact-dot--paused"></i>пауза/очередь</span>
                            <span><i class="planfact-dot planfact-dot--empty"></i>нет данных</span>
                        </div>
                    </div>
                </div>
                <div class="planfact-table">
                    <div class="planfact-table-head">
                        <span>Проект</span>
                        <span>Сроки: план продаж <span data-fact-field>/ факт</span></span>
                        <span data-fact-field>Этапы</span>
                        <span data-fact-field>Статус</span>
                        <span>План</span>
                        <span data-fact-field>Отклонение</span>
                    </div>
                    <div class="planfact-months">
                        <div class="planfact-months__timeline">
                            <div class="planfact-months__label-spacer"></div>
                            <div class="planfact-months__rail">
                                ${renderPlanFactYearBoundaries(startDate, endDate)}
                                ${renderPlanFactYearLabels()}
                                ${renderPlanFactMonthTicks(startDate, endDate)}
                            </div>
                        </div>
                    </div>
                    ${detailsRowsHtml}
                </div>
            </div>
        </details>
    `;
}

function renderPlanFactHeatmapSummary(projects, startDate, endDate) {
    const plannedProjects = projects.filter(project => project.isPlanned);
    const activeProjects = plannedProjects.filter(project => project.status.type === 'active');
    const monthlyLoadRange = getProductionProgramMonthlyLoadRange(endDate);
    const monthlyLoadMonths = getCenterMonthlyLoadMonthsForRange(monthlyLoadRange.startDate, monthlyLoadRange.endDate);

    if (monthlyLoadMonths.length === 0) {
        return '';
    }

    const workHoursByMonth = getActiveProjectActualHoursByMonth(activeProjects);
    const planHoursByMonth = getPlanProjectHoursByMonth(plannedProjects);
    const loadMonths = getMonthlyProjectLoadRows(monthlyLoadMonths, workHoursByMonth, planHoursByMonth, {
        showAllForecastHours: true
    });
    const currentMonth = getCurrentCenterMonthlyLoadRecord(loadMonths);
    const currentMonthKey = `${currentMonth.year}-${String(currentMonth.monthIndex + 1).padStart(2, '0')}`;
    const currentProjectHours = workHoursByMonth.get(currentMonthKey) || 0;
    const currentProjectPercent = currentMonth.workPercent;
    const currentPlanHours = planHoursByMonth.get(currentMonthKey) || 0;
    const currentPlanPercent = currentMonth.forecastPercent;

    return `
        <div class="center-monthly-load-card tech-monthly-load-card${getProductionProgramFactHiddenClass()}" data-fact-toggle-block="planfact-monthly-load">
            <div class="center-monthly-load-card__header">
                <div>
                    <h4 class="center-monthly-load-card__title">Общая загрузка по месяцам</h4>
                    <p class="center-monthly-load-card__subtitle">ТехБлок: плановая загрузка из «ПП» / фонд часов<span data-fact-field>. Факт показывается отдельным рядом</span>.</p>
                </div>
                <div class="center-monthly-load-card__actions">
                    <div class="center-monthly-load-card__badge" data-fact-field>${escapeHtml(currentProjectPercent === null ? '-' : `${currentProjectPercent}%`)} • ${escapeHtml(currentMonth.label)}</div>
                </div>
            </div>
            <div class="center-monthly-load-kpis">
                ${renderCenterMonthlyLoadKpi('Факт часов', formatHoursValue(currentProjectHours), 'ч', 'green', 'data-fact-field')}
                ${renderCenterMonthlyLoadKpi('Фонд часов', formatHoursValue(currentMonth.totalHours), 'ч', 'blue')}
                ${renderCenterMonthlyLoadKpi('В работе', currentProjectPercent === null ? '-' : `${currentProjectPercent}%`, '', 'green', 'data-fact-field')}
                ${renderCenterMonthlyLoadKpi('План', currentPlanPercent === null ? '-' : `${currentPlanPercent}%`, `${currentMonth.label} • ${formatHoursValue(currentPlanHours)} ч`, 'blue')}
            </div>
            <div class="center-monthly-load-heatwrap">
                <div class="center-monthly-load-heat-title">
                    Общая загрузка
                    <small><span data-fact-field>в работе и </span>будущая потребность</small>
                </div>
                ${renderMonthlyLoadScale(loadMonths, { forecastLabel: 'план' })}
            </div>
        </div>
    `;
}

function getProductionProgramMonthlyLoadRange(endDate) {
    const referenceDate = getReferenceDate();
    const startDate = new Date(referenceDate.getFullYear(), 0, 1);
    const fallbackEndDate = new Date(referenceDate.getFullYear() + 1, referenceDate.getMonth() + 1, 1);

    return {
        startDate,
        endDate: endDate instanceof Date && !isNaN(endDate.getTime()) && endDate > startDate
            ? endDate
            : fallbackEndDate
    };
}

function getPlanProjectHoursByMonth(projects = null) {
    const result = new Map();
    const sourceProjects = projects || getTechBlockProjectPlanFactDataFromTables();

    sourceProjects.forEach(project => {
        (project.planHoursByCenterMonth || new Map()).forEach(monthMap => {
            if (!(monthMap instanceof Map)) return;
            monthMap.forEach((value, monthKey) => addMonthMapValue(result, monthKey, value));
        });
    });

    return result;
}

function renderPlanFactHeatmapMonthLabels(startDate, endDate) {
    return getPlanFactDisplayTicks().map(tick => {
        const date = new Date(tick.year, tick.month, 1);
        const nextDate = new Date(tick.year, tick.month + 1, 1);
        const startPosition = getDatePositionPercent(date, startDate, endDate);
        const endPosition = getDatePositionPercent(nextDate, startDate, endDate);
        const position = Math.min(100, startPosition + Math.max(0, endPosition - startPosition) / 2);
        const edgeClass = position <= 0.5
            ? ' planfact-heatmap-labels__month--start'
            : position >= 99.5
                ? ' planfact-heatmap-labels__month--end'
                : '';

        return `
            <span class="planfact-heatmap-labels__month${edgeClass}" style="left: ${position}%;">
                ${escapeHtml(formatPlanFactMonthLabel(date))}
            </span>
        `;
    }).join('');
}

function renderPlanFactHeatmapYearLabels() {
    return getPlanFactYearSegments().map(segment => `
        <span class="planfact-heatmap-labels__year" style="left: ${segment.start + segment.width / 2}%;">
            ${segment.year}
        </span>
    `).join('');
}

function getProjectHeatmapData(projects, startDate, endDate) {
    const monthIntervals = getPlanFactMonthIntervals(startDate, endDate);
    const projectRanges = projects.map(project => getProjectHeatmapRange(project));
    const monthCounts = getOutsidePlanMonthCounts(projectRanges, monthIntervals);

    return {
        monthCounts,
        maxMonthlyCount: Math.max(0, ...monthCounts.map(month => month.count))
    };
}

function getProjectHoursHeatmapData(projects, startDate, endDate) {
    const monthIntervals = getPlanFactMonthIntervals(startDate, endDate);
    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);
    const planByMonth = new Map();
    const actualByMonth = new Map();
    const capacityByMonth = new Map();

    projects.forEach(project => {
        (project.planHoursByCenterMonth || new Map()).forEach(monthMap => {
            monthMap.forEach((value, monthKey) => addMonthMapValue(planByMonth, monthKey, value));
        });
        (project.actualHoursByCenterMonth || new Map()).forEach(monthMap => {
            monthMap.forEach((value, monthKey) => addMonthMapValue(actualByMonth, monthKey, value));
        });
    });

    capacityByCenterMonth.forEach(monthMap => {
        monthMap.forEach((value, monthKey) => addMonthMapValue(capacityByMonth, monthKey, value.availableHours || 0));
    });

    const monthCounts = monthIntervals.map(month => {
        const monthKey = getMonthKey(month.start);
        const availableHours = capacityByMonth.get(monthKey) || 0;
        const actualHours = actualByMonth.get(monthKey) || 0;
        const planHours = planByMonth.get(monthKey) || 0;

        return {
            ...month,
            monthKey,
            planHours,
            actualHours,
            availableHours,
            loadPercent: availableHours > 0 ? Math.round((actualHours / availableHours) * 100) : null,
            planPercent: availableHours > 0 ? Math.round((planHours / availableHours) * 100) : null,
            count: actualHours
        };
    });

    return {
        monthCounts,
        maxMonthlyCount: 100
    };
}

function getProjectCenterLoadData(projects) {
    const centerOrder = ['ЦНГО', 'ЦДЦМ', 'ЦДПМ', 'ЦОТ', 'ЛОИ', 'УВП'];
    const centerMap = new Map();
    const outsideProjects = projects.filter(project => !project.isPlanned);
    const { startDate, endDate } = getPlanFactDisplayRange();
    const monthIntervals = getPlanFactMonthIntervals(startDate, endDate);
    const capacityByCenterMonth = getTechBlockCenterCapacityByMonth(monthIntervals);

    projects.filter(project => project.isPlanned).forEach(project => {
        const centers = getProjectCentersForLoad(project);

        if (centers.length === 0) return;

        centers.forEach(center => {
            const planShare = getProjectPlanHoursForCenter(project, center);
            const actualShare = getProjectActualHoursForCenter(project, center);

            if (!centerMap.has(center)) {
                centerMap.set(center, {
                    key: center,
                    fullName: getProjectStageFullLabel(center),
                    projectIds: new Set(),
                    projectRanges: [],
                    planMonthHours: new Map(),
                    actualMonthHours: new Map(),
                    planHours: 0,
                    actualHours: 0,
                    planStart: null,
                    planEnd: null,
                    factStart: null,
                    factEnd: null
                });
            }

            const row = centerMap.get(center);
            row.projectIds.add(project.id);
            row.projectRanges.push(getProjectHeatmapRange(project));
            row.planHours += planShare;
            row.actualHours += actualShare;
            mergeNestedMonthMap(row.planMonthHours, project.planHoursByCenterMonth, center);
            mergeNestedMonthMap(row.actualMonthHours, project.actualHoursByCenterMonth, center);
            row.planStart = getEarlierDate(row.planStart, project.planStart);
            row.planEnd = getLaterDate(row.planEnd, project.planEnd || project.planStart);
            row.factStart = getEarlierDate(row.factStart, project.factStart);
            row.factEnd = getLaterDate(row.factEnd, project.factEnd || project.factStart);
        });
    });

    const centers = Array.from(centerMap.values())
        .filter(center => center.projectIds.size > 0 || center.planHours > 0 || center.actualHours > 0)
        .map(center => ({
            ...center,
            projectCount: center.projectIds.size,
            loadPercent: center.planHours > 0 ? Math.round((center.actualHours / center.planHours) * 100) : null,
            tone: getCenterLoadTone(center.planHours, center.actualHours),
            monthLoads: getCenterMonthLoadCells(center.planMonthHours, center.actualMonthHours, capacityByCenterMonth.get(center.key), monthIntervals)
        }))
        .sort((a, b) => {
            const aIndex = centerOrder.indexOf(a.key);
            const bIndex = centerOrder.indexOf(b.key);
            const aRank = aIndex === -1 ? centerOrder.length : aIndex;
            const bRank = bIndex === -1 ? centerOrder.length : bIndex;
            if (aRank !== bRank) return aRank - bRank;
            return a.key.localeCompare(b.key, 'ru');
        });
    const totalPlanHours = centers.reduce((sum, center) => sum + center.planHours, 0);
    const totalActualHours = centers.reduce((sum, center) => sum + center.actualHours, 0);
    const projectIds = new Set();
    centers.forEach(center => center.projectIds.forEach(id => projectIds.add(id)));
    const outsidePlan = getOutsidePlanProjectLoadData(outsideProjects, centerOrder, capacityByCenterMonth, monthIntervals);

    return {
        centers,
        totalPlanHours,
        totalActualHours,
        executionPercent: totalPlanHours > 0 ? Math.round((totalActualHours / totalPlanHours) * 100) : null,
        overloadedCount: centers.filter(center => center.tone === 'over').length,
        riskCount: centers.filter(center => center.tone === 'risk').length,
        projectCount: projectIds.size,
        maxMonthlyCount: 100,
        outsidePlan
    };
}

function getProjectCentersForLoad(project) {
    const stageCenters = (project.centerStages || [])
        .map(stage => stage.shortLabel)
        .filter(Boolean);
    const planCenters = project.planHoursByCenter instanceof Map ? Array.from(project.planHoursByCenter.keys()) : [];
    const actualCenters = project.actualHoursByCenter instanceof Map ? Array.from(project.actualHoursByCenter.keys()) : [];

    return Array.from(new Set([...stageCenters, ...planCenters, ...actualCenters]));
}

function getCenterMonthLoadCells(planMonthHours, actualMonthHours, capacityMonthMap, monthIntervals) {
    return monthIntervals.map(month => {
        const monthKey = getMonthKey(month.start);
        const planHours = planMonthHours instanceof Map ? (planMonthHours.get(monthKey) || 0) : 0;
        const actualHours = actualMonthHours instanceof Map ? (actualMonthHours.get(monthKey) || 0) : 0;
        const capacity = capacityMonthMap instanceof Map ? capacityMonthMap.get(monthKey) : null;
        const availableHours = capacity ? capacity.availableHours || 0 : 0;
        const loadPercent = availableHours > 0 ? Math.round((actualHours / availableHours) * 100) : null;
        const planPercent = availableHours > 0 ? Math.round((planHours / availableHours) * 100) : null;

        return {
            ...month,
            monthKey,
            planHours,
            actualHours,
            availableHours,
            absenceHours: capacity ? capacity.absenceHours || 0 : 0,
            grossHours: capacity ? capacity.grossHours || 0 : 0,
            loadPercent,
            planPercent,
            count: actualHours
        };
    });
}

function getProjectHoursForCenter(hoursByCenter, center, centerCount, totalHours) {
    if (hoursByCenter instanceof Map && hoursByCenter.has(center)) {
        return hoursByCenter.get(center) || 0;
    }

    return totalHours > 0 && centerCount > 0 ? totalHours / centerCount : 0;
}

function getProjectPlanHoursForCenter(project, center) {
    const hoursByCenter = project.planHoursByCenter;
    if (!(hoursByCenter instanceof Map)) return 0;
    return hoursByCenter.get(center) || 0;
}

function getProjectActualHoursForCenter(project, center) {
    const hoursByCenter = project.actualHoursByCenter;
    if (!(hoursByCenter instanceof Map)) return 0;
    return hoursByCenter.get(center) || 0;
}

function getOutsidePlanProjectLoadData(projects, centerOrder, capacityByCenterMonth, monthIntervals) {
    const centerMap = new Map();
    let totalActualHours = 0;
    let totalPlanHours = 0;

    projects.forEach(project => {
        const centers = getProjectCentersForLoad(project);
        const centerKeys = centers.length > 0 ? centers : getProjectManagementFallbackCenters(project);

        totalActualHours += project.actualHours || 0;
        totalPlanHours += project.planHours || 0;

        centerKeys.forEach(center => {
            const planShare = getProjectPlanHoursForCenter(project, center);
            const actualShare = getProjectHoursForCenter(project.actualHoursByCenter, center, centerKeys.length, project.actualHours);

            if (!centerMap.has(center)) {
                centerMap.set(center, {
                    key: center,
                    fullName: getProjectStageFullLabel(center),
                    projectIds: new Set(),
                    projectRanges: [],
                    planMonthHours: new Map(),
                    actualMonthHours: new Map(),
                    statusMap: new Map(),
                    planHours: 0,
                    actualHours: 0,
                    factStart: null,
                    factEnd: null
                });
            }

            const row = centerMap.get(center);
            row.projectIds.add(project.id);
            addProjectStatusCount(row.statusMap, project.status);
            row.planHours += planShare;
            if (project.status.type === 'active') {
                const heatmapRange = getProjectFactHeatmapRange(project);
                row.projectRanges.push({
                    id: project.id,
                    start: heatmapRange.start,
                    end: heatmapRange.end
                });
                row.factStart = getEarlierDate(row.factStart, heatmapRange.start);
                row.factEnd = getLaterDate(row.factEnd, heatmapRange.end || heatmapRange.start);
            }
            row.actualHours += actualShare;
            if (project.status.type === 'active') {
                mergeNestedMonthMap(row.planMonthHours, project.planHoursByCenterMonth, center);
                mergeNestedMonthMap(row.actualMonthHours, project.actualHoursByCenterMonth, center);
            }
        });
    });

    const centers = Array.from(centerMap.values())
        .map(center => ({
            ...center,
            projectCount: center.projectIds.size,
            statusSummary: getProjectStatusSummaryFromMap(center.statusMap),
            monthCounts: getCenterMonthLoadCells(center.planMonthHours, center.actualMonthHours, capacityByCenterMonth.get(center.key), monthIntervals)
        }))
        .sort((a, b) => {
            const aIndex = centerOrder.indexOf(a.key);
            const bIndex = centerOrder.indexOf(b.key);
            const aRank = aIndex === -1 ? centerOrder.length : aIndex;
            const bRank = bIndex === -1 ? centerOrder.length : bIndex;
            if (aRank !== bRank) return aRank - bRank;
            return a.key.localeCompare(b.key, 'ru');
        });

    return {
        projects,
        centers,
        totalPlanHours,
        totalActualHours,
        projectCount: projects.length,
        maxMonthlyCount: 100
    };
}

function getProjectManagementFallbackCenters(project) {
    const fallbackCenters = (project.stages || [])
        .filter(stage => (
            (stage.shortLabel === 'ЛОИ' || stage.shortLabel === 'УВП')
            && stage.status
            && stage.status.type !== 'empty'
        ))
        .map(stage => stage.shortLabel);

    return fallbackCenters.length > 0 ? fallbackCenters : ['ЛОИ', 'УВП'];
}

function getProjectHeatmapRange(project) {
    const factRange = getProjectFactHeatmapRange(project);
    if (factRange.start || factRange.end) return factRange;

    return {
        id: project.id,
        start: project.planStart,
        end: project.planEnd || project.planStart
    };
}

function getProjectFactHeatmapRange(project) {
    return {
        id: project.id,
        start: project.factStart,
        end: project.factEnd || project.factStart
    };
}

function getProjectStatusSummary(projects) {
    const statusMap = new Map();

    projects.forEach(project => {
        addProjectStatusCount(statusMap, project.status);
    });

    return getProjectStatusSummaryFromMap(statusMap);
}

function addProjectStatusCount(statusMap, statusValue) {
    const status = statusValue || { label: '-', type: 'empty' };
    const key = status.type || 'empty';

    if (!statusMap.has(key)) {
        statusMap.set(key, {
            type: key,
            label: status.label || '-',
            count: 0
        });
    }

    statusMap.get(key).count += 1;
}

function getProjectStatusSummaryFromMap(statusMap) {
    const statusOrder = ['active', 'done', 'paused', 'cancelled', 'empty'];

    return Array.from(statusMap.values()).sort((a, b) => {
        const aIndex = statusOrder.indexOf(a.type);
        const bIndex = statusOrder.indexOf(b.type);
        const aRank = aIndex === -1 ? statusOrder.length : aIndex;
        const bRank = bIndex === -1 ? statusOrder.length : bIndex;
        if (aRank !== bRank) return aRank - bRank;
        return a.label.localeCompare(b.label, 'ru');
    });
}

function getPlanFactMonthIntervals(startDate, endDate) {
    const months = [];
    const cursor = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
    const last = new Date(endDate.getFullYear(), endDate.getMonth(), 1);

    while (cursor < last) {
        const monthStart = new Date(cursor.getFullYear(), cursor.getMonth(), 1);
        const monthEnd = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
        months.push({
            start: monthStart,
            end: monthEnd,
            label: `${formatMonthName(monthStart.getMonth())} ${monthStart.getFullYear()}`,
            left: getDatePositionPercent(monthStart, startDate, endDate),
            width: Math.max(1.5, getDatePositionPercent(monthEnd, startDate, endDate) - getDatePositionPercent(monthStart, startDate, endDate))
        });
        cursor.setMonth(cursor.getMonth() + 1);
    }

    return months;
}

function getMonthKey(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '';
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function parseMonthKeyToDate(monthKey) {
    if (typeof monthKey !== 'string') return null;
    const match = monthKey.match(/^(\d{4})-(\d{2})$/);
    if (!match) return null;

    const year = Number(match[1]);
    const monthIndex = Number(match[2]) - 1;
    if (!Number.isFinite(year) || !Number.isFinite(monthIndex) || monthIndex < 0 || monthIndex > 11) {
        return null;
    }

    return new Date(year, monthIndex, 1);
}

function getMonthKeysBetween(start, end) {
    const rangeStart = start || end;
    const rangeEnd = end || start;
    if (!(rangeStart instanceof Date) || isNaN(rangeStart.getTime())) return [];

    const first = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);
    const last = rangeEnd instanceof Date && !isNaN(rangeEnd.getTime())
        ? new Date(rangeEnd.getFullYear(), rangeEnd.getMonth(), 1)
        : first;
    const keys = [];
    const cursor = new Date(first);

    while (cursor <= last) {
        keys.push(getMonthKey(cursor));
        cursor.setMonth(cursor.getMonth() + 1);
    }

    return keys;
}

function getOutsidePlanMonthCounts(projectRanges, monthIntervals) {
    return monthIntervals.map(month => ({
        ...month,
        count: projectRanges.filter(project => (
            project && doesDateRangeIntersectPeriod(project.start, project.end, month.start, month.end)
        )).length
    }));
}

function doesDateRangeIntersectPeriod(start, end, periodStart, periodEnd) {
    const rangeStart = start || end;
    const rangeEnd = end || start;

    if (!(rangeStart instanceof Date) || isNaN(rangeStart.getTime())) return false;

    return rangeStart < periodEnd && rangeEnd >= periodStart;
}

function getEarlierDate(current, candidate) {
    if (!(candidate instanceof Date) || isNaN(candidate.getTime())) return current;
    if (!(current instanceof Date) || isNaN(current.getTime())) return candidate;
    return candidate < current ? candidate : current;
}

function getLaterDate(current, candidate) {
    if (!(candidate instanceof Date) || isNaN(candidate.getTime())) return current;
    if (!(current instanceof Date) || isNaN(current.getTime())) return candidate;
    return candidate > current ? candidate : current;
}

function getCenterLoadTone(planHours, actualHours) {
    if (planHours <= 0 && actualHours <= 0) return 'neutral';
    if (planHours <= 0 && actualHours > 0) return 'over';

    const ratio = actualHours / planHours;
    if (ratio > 1.05) return 'over';
    if (ratio < 0.75) return 'risk';
    return 'ok';
}

function renderProjectCenterLoadBlock(data, startDate, endDate) {
    if (!data || data.centers.length === 0) {
        return '';
    }

    return `
        <section class="centerload-block${getProductionProgramFactHiddenClass()}" data-fact-toggle-block="project-center-load">
            <div class="centerload-header">
                <div>
                    <div class="centerload-title">Проекты из плана: сроки и часы по центрам</div>
                    <div class="centerload-subtitle">Плановые сроки и часы из «ПП»<span data-fact-field> в сравнении с фактическим выполнением из «Загрузка_ТехБлок»</span>.</div>
                </div>
                <div class="centerload-header__actions">
                    <div class="centerload-period">2025-2027<br>план <span data-fact-field>/ факт</span></div>
                </div>
            </div>
            <div class="centerload-kpis">
                ${renderCenterLoadKpi('Плановые часы', formatHoursValue(data.totalPlanHours), 'ч', 'blue')}
                ${renderCenterLoadKpi('Фактические часы', formatHoursValue(data.totalActualHours), 'ч', 'green', 'data-fact-field')}
                ${renderCenterLoadKpi('Исполнение плана', data.executionPercent === null ? '-' : `${data.executionPercent}%`, '', getExecutionPercentTone(data.executionPercent), 'data-fact-field')}
                ${renderCenterLoadKpi('Перегрузка', data.overloadedCount, formatCenterWord(data.overloadedCount), 'red', 'data-fact-field')}
                ${renderCenterLoadKpi('Проектов в плане', data.projectCount, formatProjectWord(data.projectCount), 'neutral')}
            </div>
            <div class="centerload-legend">
                <span><i class="centerload-dot centerload-dot--plan"></i>плановые сроки</span>
                <span data-fact-field><i class="centerload-heatline" aria-hidden="true"></i>факт: часы / доступная мощность</span>
                <span data-fact-field><i class="centerload-dot centerload-dot--low"></i>до 70%</span>
                <span data-fact-field><i class="centerload-dot centerload-dot--high"></i>свыше 100%</span>
            </div>
            <div class="centerload-table">
                <div class="centerload-row centerload-row--head">
                    <span>Центр</span>
                    <span>Сроки: план <span data-fact-field>/ факт</span></span>
                    <span>План</span>
                    <span data-fact-field>Факт</span>
                    <span data-fact-field>Отклонение</span>
                    <span data-fact-field>Загрузка</span>
                </div>
                <div class="centerload-calendar">
                    <div class="centerload-calendar__timeline">
                        <div class="centerload-calendar__label-spacer"></div>
                        <div class="centerload-calendar__rail">
                            ${renderPlanFactYearBoundaries(startDate, endDate)}
                            ${renderCenterLoadYearLabels()}
                            ${renderCenterLoadMonthTicks(startDate, endDate)}
                        </div>
                    </div>
                </div>
                ${data.centers.map(center => renderCenterLoadRow(center, startDate, endDate, data.maxMonthlyCount)).join('')}
            </div>
        </section>
    `;
}

function renderOutsidePlanLoadImpactBlock(data) {
    const rows = getOutsidePlanLoadImpactRows(data);

    if (rows.length === 0) {
        return '';
    }

    const totalOutsideHours = rows.reduce((sum, row) => sum + row.outsideHours, 0);
    const totalPlanHours = rows.reduce((sum, row) => sum + row.planHours, 0);
    const totalImpact = totalPlanHours > 0 ? Math.round((totalOutsideHours / totalPlanHours) * 100) : null;

    return `
        <div class="outside-impact-block">
            <div class="outside-impact-block__header">
                <div>
                    <div class="outside-impact-block__title">Внеплановая нагрузка</div>
                    <div class="outside-impact-block__subtitle">Фактические часы проектов вне плана относительно плановых часов центра.</div>
                </div>
                <div class="outside-impact-block__total">
                    <span>${formatHoursValue(totalOutsideHours)} ч</span>
                    <strong>${totalImpact === null ? 'нет плана' : `${totalImpact}%`}</strong>
                </div>
            </div>
            <div class="outside-impact-list">
                ${rows.map(row => renderOutsidePlanLoadImpactRow(row)).join('')}
            </div>
        </div>
    `;
}

function getOutsidePlanLoadImpactRows(data) {
    const planMap = new Map((data.centers || []).map(center => [center.key, center]));
    const outsideCenters = data.outsidePlan && data.outsidePlan.centers ? data.outsidePlan.centers : [];
    const outsideMap = new Map(outsideCenters.map(center => [center.key, center]));
    const keys = Array.from(new Set([...planMap.keys(), ...outsideMap.keys()]));
    const maxHours = Math.max(
        1,
        ...keys.map(key => Math.max(
            planMap.get(key)?.planHours || 0,
            outsideMap.get(key)?.actualHours || 0
        ))
    );

    return keys
        .map(key => {
            const plan = planMap.get(key);
            const outside = outsideMap.get(key);
            const planHours = plan?.planHours || 0;
            const outsideHours = outside?.actualHours || 0;
            const impactPercent = planHours > 0 ? Math.round((outsideHours / planHours) * 100) : null;

            return {
                key,
                fullName: getProjectStageFullLabel(key),
                planHours,
                outsideHours,
                impactPercent,
                planWidth: Math.max(0, Math.min(100, (planHours / maxHours) * 100)),
                outsideWidth: Math.max(0, Math.min(100, (outsideHours / maxHours) * 100)),
                tone: getOutsideImpactTone(impactPercent, planHours, outsideHours)
            };
        })
        .filter(row => row.planHours > 0 || row.outsideHours > 0);
}

function getOutsideImpactTone(percent, planHours, outsideHours) {
    if (outsideHours <= 0) return 'neutral';
    if (planHours <= 0) return 'over';
    if (percent >= 50) return 'over';
    if (percent >= 20) return 'risk';
    return 'ok';
}

function renderOutsidePlanLoadImpactRow(row) {
    const impactText = row.impactPercent === null ? 'нет плана' : `${row.impactPercent}%`;
    const planText = `${formatHoursValue(row.planHours)} ч`;
    const outsideText = `${formatHoursValue(row.outsideHours)} ч`;

    return `
        <div class="outside-impact-row outside-impact-row--${row.tone}">
            <div class="outside-impact-row__center">
                <span title="${escapeHtml(row.key)}">${escapeHtml(row.fullName)}</span>
                <small>${escapeHtml(row.key)}</small>
            </div>
            <div class="outside-impact-row__compare" aria-label="Плановые и внеплановые часы">
                <div class="outside-impact-row__hours outside-impact-row__hours--plan">${escapeHtml(planText)}</div>
                <div class="outside-impact-row__bars">
                    <div class="outside-impact-bar outside-impact-bar--plan" style="width: ${row.planWidth}%;"></div>
                    <div class="outside-impact-bar outside-impact-bar--outside" style="width: ${row.outsideWidth}%;"></div>
                </div>
                <div class="outside-impact-row__hours outside-impact-row__hours--outside">${escapeHtml(outsideText)}</div>
            </div>
            <div class="outside-impact-row__value">${escapeHtml(impactText)}</div>
        </div>
    `;
}

function renderOutsidePlanHeatmap(monthCounts, maxMonthlyCount, tooltipSuffix = 'в работе') {
    if (!monthCounts || monthCounts.length === 0) {
        return '<div class="outside-plan-heatmap outside-plan-heatmap--empty">-</div>';
    }

    return `
        <div class="outside-plan-heatmap" aria-label="Распределение проектов по месяцам">
            <div class="outside-plan-heatmap__track"></div>
            ${monthCounts.map(month => renderHeatmapMonthCell(month, maxMonthlyCount, tooltipSuffix)).join('')}
        </div>
    `;
}

function renderHeatmapMonthCell(month, maxMonthlyCount, tooltipSuffix) {
    if (month.loadPercent !== undefined) {
        const hasHours = (month.actualHours || 0) > 0 || (month.planHours || 0) > 0;
        const percentText = month.loadPercent === null ? '-' : `${month.loadPercent}%`;
        const title = [
            `${month.label}`,
            `план ${formatHoursValue(month.planHours || 0)} ч`,
            `факт ${formatHoursValue(month.actualHours || 0)} ч`,
            `доступно ${formatHoursValue(month.availableHours || 0)} ч`,
            `загрузка ${percentText}`
        ].join(' / ');

        return `
            <span class="outside-plan-heatmap__cell"
                  title="${escapeHtml(title)}"
                  style="left: ${month.left}%; width: ${month.width}%; background: ${getLoadHeatColor(month.loadPercent, hasHours)};">
                ${hasHours && month.loadPercent !== null ? `<span>${percentText}</span>` : ''}
            </span>
        `;
    }

    return `
        <span class="outside-plan-heatmap__cell"
              title="${escapeHtml(month.label)}: ${month.count} ${formatProjectWord(month.count)}${tooltipSuffix ? ` ${escapeHtml(tooltipSuffix)}` : ''}"
              style="left: ${month.left}%; width: ${month.width}%; background: ${getOutsidePlanHeatColor(month.count, maxMonthlyCount)};">
            ${month.count > 0 ? `<span>${month.count}</span>` : ''}
        </span>
    `;
}

function getLoadHeatColor(percent, hasHours) {
    if (!hasHours) return '#edf2f7';
    if (percent === null || percent === undefined) return '#dbe3ec';
    if (percent <= 70) return '#69ab81';
    if (percent <= 90) return mixHexColor('#69ab81', '#d7a44b', (percent - 70) / 20);
    if (percent <= 100) return mixHexColor('#d7a44b', '#d96d61', (percent - 90) / 10);
    return '#d96d61';
}

function getForecastLoadColor(percent, hasHours) {
    if (!hasHours) return '#edf2f7';
    if (percent === null || percent === undefined) return '#dbe3ec';
    if (percent <= 80) return '#5b9bd5';
    if (percent <= 100) return mixHexColor('#5b9bd5', '#dc8a34', (percent - 80) / 20);
    return '#dc8a34';
}

function getOutsidePlanHeatColor(count, maxCount) {
    if (!count || count <= 0 || !maxCount) return '#edf2f7';

    const ratio = Math.max(0, Math.min(1, count / maxCount));
    if (ratio < 0.5) {
        const t = ratio / 0.5;
        return mixHexColor('#69ab81', '#d7a44b', t);
    }

    return mixHexColor('#d7a44b', '#d96d61', (ratio - 0.5) / 0.5);
}

function mixHexColor(from, to, amount) {
    const start = hexToRgb(from);
    const end = hexToRgb(to);
    const rgb = start.map((value, index) => Math.round(value + (end[index] - value) * amount));
    return `rgb(${rgb.join(', ')})`;
}

function hexToRgb(hex) {
    const value = hex.replace('#', '');
    return [
        parseInt(value.slice(0, 2), 16),
        parseInt(value.slice(2, 4), 16),
        parseInt(value.slice(4, 6), 16)
    ];
}

function renderCenterLoadYearLabels() {
    return getPlanFactYearSegments().map(segment => `
        <span class="centerload-year-label" style="left: ${segment.start + segment.width / 2}%;">
            ${segment.year}
        </span>
    `).join('');
}

function renderCenterLoadMonthTicks(startDate, endDate) {
    return getPlanFactDisplayTicks().map(tick => {
        const date = new Date(tick.year, tick.month, 1);
        const position = getDatePositionPercent(date, startDate, endDate);
        const edgeClass = position <= 0.5
            ? ' centerload-month-tick--start'
            : position >= 99.5
                ? ' centerload-month-tick--end'
                : '';

        return `
            <span class="centerload-month-tick${edgeClass}" style="left: ${position}%;">
                ${escapeHtml(formatPlanFactMonthLabel(date))}
            </span>
        `;
    }).join('');
}

function renderCenterLoadKpi(title, value, note, tone, extraAttributes = '') {
    return `
        <div class="centerload-kpi centerload-kpi--${tone}" ${extraAttributes}>
            <span class="centerload-kpi__title">${escapeHtml(title)}</span>
            <span class="centerload-kpi__value">${escapeHtml(String(value))}</span>
            ${note ? `<span class="centerload-kpi__note">${escapeHtml(String(note))}</span>` : ''}
        </div>
    `;
}

function getExecutionPercentTone(percent) {
    if (percent === null || percent === undefined) return 'neutral';
    if (percent > 105) return 'red';
    if (percent < 75) return 'amber';
    return 'green';
}

function renderCenterLoadRow(center, startDate, endDate, maxMonthlyCount) {
    const deviation = Math.round(center.actualHours - center.planHours);
    const deviationTone = deviation > 0 ? 'over' : deviation < 0 ? 'under' : 'neutral';
    const deviationArrow = deviation > 0 ? '▲' : deviation < 0 ? '▼' : '•';
    const deviationText = formatHoursValue(Math.abs(deviation));
    const loadText = center.loadPercent === null ? '-' : `${center.loadPercent}%`;

    return `
        <div class="centerload-row centerload-row--${center.tone}">
            <div class="centerload-center">
                <div class="centerload-center__name" title="${escapeHtml(center.key)}">${escapeHtml(center.fullName)}</div>
                <div class="centerload-center__meta">${escapeHtml(center.key)} · ${center.projectCount} ${formatProjectWord(center.projectCount)}</div>
            </div>
            <div class="centerload-timeline">
                ${renderCenterLoadTimelineLine('план', center.planStart, center.planEnd, startDate, endDate, 'plan')}
                ${renderCenterLoadHeatmapLine('факт', center.monthLoads, maxMonthlyCount)}
            </div>
            <div class="centerload-metric" data-label="План">${formatHoursValue(center.planHours)} <small>ч</small></div>
            <div class="centerload-metric" data-label="Факт" data-fact-field>${formatHoursValue(center.actualHours)} <small>ч</small></div>
            <div class="centerload-metric centerload-deviation centerload-deviation--${deviationTone}" data-label="Отклонение" data-fact-field>
                <span class="centerload-deviation__arrow">${deviationArrow}</span>
                <span>${deviationText}</span> <small>ч</small>
            </div>
            <div class="centerload-load-cell" data-label="Загрузка" data-fact-field><span class="centerload-badge centerload-badge--${center.tone}">${escapeHtml(loadText)}</span></div>
        </div>
    `;
}

function renderCenterLoadHeatmapLine(label, monthCounts, maxMonthlyCount) {
    return `
        <div class="centerload-timeline-line centerload-timeline-line--heat" data-fact-field>
            <span class="centerload-timeline-line__label">${escapeHtml(label)}</span>
            <div class="centerload-timeline-line__rail">
                ${renderOutsidePlanHeatmap(monthCounts, maxMonthlyCount, 'из плана')}
            </div>
        </div>
    `;
}

function renderCenterLoadTimelineLine(label, start, end, rangeStart, rangeEnd, type) {
    const hasStart = start instanceof Date && !isNaN(start.getTime());
    const hasEnd = end instanceof Date && !isNaN(end.getTime());

    if (!hasStart && !hasEnd) {
        return `
            <div class="centerload-timeline-line centerload-timeline-line--${type}">
                <span class="centerload-timeline-line__label">${label}</span>
                <div class="centerload-timeline-line__rail">
                    ${renderPlanFactYearBoundaries(rangeStart, rangeEnd)}
                    <span class="centerload-timeline__empty">-</span>
                </div>
            </div>
        `;
    }

    const lineStart = hasStart ? start : end;
    const lineEnd = hasEnd ? end : start;
    const startPosition = getDatePositionPercent(lineStart, rangeStart, rangeEnd);
    const endPosition = getDatePositionPercent(lineEnd, rangeStart, rangeEnd);
    const left = Math.min(startPosition, endPosition);
    const width = Math.max(2, Math.abs(endPosition - startPosition));
    const dateText = formatTimelineRangeLabel(lineStart, lineEnd);

    return `
        <div class="centerload-timeline-line centerload-timeline-line--${type}">
            <span class="centerload-timeline-line__label">${label}</span>
            <div class="centerload-timeline-line__rail">
                ${renderPlanFactYearBoundaries(rangeStart, rangeEnd)}
                <div class="centerload-timeline__track"></div>
                <div class="centerload-timeline__progress centerload-timeline__progress--${type}" style="left: ${left}%; width: ${width}%;"></div>
                <span class="centerload-timeline__dates" style="left: ${left}%; width: ${width}%;">${escapeHtml(dateText)}</span>
            </div>
        </div>
    `;
}

function formatHoursValue(value) {
    const rounded = Math.round(value || 0);
    return Object.is(rounded, -0) ? '0' : rounded.toLocaleString('ru-RU');
}

function formatTimelineRangeLabel(start, end) {
    const startText = formatDateShort(start);
    const endText = formatDateShort(end);
    return startText === endText ? startText : `${startText} - ${endText}`;
}

function formatDateShort(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '-';
    return date.toLocaleDateString('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: '2-digit'
    });
}

function formatDateFull(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return '-';
    return date.toLocaleDateString('ru-RU', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
    });
}

function renderPlanFactSummaryCard(title, value, subtitle, tone, extraAttributes = '') {
    return `
        <div class="planfact-summary-card planfact-summary-card--${tone}" ${extraAttributes}>
            <span class="planfact-summary-card__title">${escapeHtml(title)}</span>
            <span class="planfact-summary-card__value">${escapeHtml(String(value))}</span>
            <span class="planfact-summary-card__subtitle">${escapeHtml(subtitle)}</span>
        </div>
    `;
}

function getTechBlockProjectTimelineRange(projects) {
    const { startDate, endDate } = getPlanFactDisplayRange();
    return { startDate, endDate };
}

function getPlanFactYearSegments() {
    return [
        { year: 2025, start: 0, width: 25, startDate: new Date(2025, 6, 1), endDate: new Date(2026, 0, 1) },
        { year: 2026, start: 25, width: 50, startDate: new Date(2026, 0, 1), endDate: new Date(2027, 0, 1) },
        { year: 2027, start: 75, width: 25, startDate: new Date(2027, 0, 1), endDate: new Date(2027, 6, 1) }
    ];
}

function getDatePositionPercent(date, startDate, endDate) {
    if (!(date instanceof Date) || isNaN(date.getTime())) return 0;

    const segments = getPlanFactYearSegments();
    const segment = segments.find(item => date >= item.startDate && date <= item.endDate);

    if (segment) {
        const segmentDuration = segment.endDate.getTime() - segment.startDate.getTime();
        const segmentProgress = segmentDuration > 0
            ? (date.getTime() - segment.startDate.getTime()) / segmentDuration
            : 0;
        const value = segment.start + Math.max(0, Math.min(1, segmentProgress)) * segment.width;
        return Math.max(0, Math.min(100, value));
    }

    if (date < segments[0].startDate) return 0;
    if (date > segments[segments.length - 1].endDate) return 100;

    const total = endDate.getTime() - startDate.getTime();
    if (total <= 0) return 0;

    const value = ((date.getTime() - startDate.getTime()) / total) * 100;
    return Math.max(0, Math.min(100, value));
}

function renderPlanFactMonthTicks(startDate, endDate) {
    const ticksData = getPlanFactDisplayTicks();
    const ticks = ticksData.map(tick => {
        const date = new Date(tick.year, tick.month, 1);
        const position = getDatePositionPercent(date, startDate, endDate);
        const edgeClass = position <= 0.5
            ? ' planfact-month-tick--start'
            : position >= 99.5
                ? ' planfact-month-tick--end'
                : '';
        return `
            <span class="planfact-month-tick${edgeClass}" style="left: ${position}%;">
                ${escapeHtml(formatPlanFactMonthLabel(date))}
            </span>
        `;
    });

    return ticks.join('');
}

function renderPlanFactYearLabels() {
    return getPlanFactYearSegments().map(segment => `
        <span class="planfact-year-label" style="left: ${segment.start + segment.width / 2}%;">
            ${segment.year}
        </span>
    `).join('');
}

function renderPlanFactYearBoundaries(startDate, endDate) {
    return [2026, 2027].map(year => {
        const date = new Date(year, 0, 1);
        const position = getDatePositionPercent(date, startDate, endDate);
        return `<span class="planfact-year-boundary" style="left: ${position}%;" aria-hidden="true"></span>`;
    }).join('');
}

function getPlanFactDisplayTicks() {
    return [
        { year: 2025, month: 6 },
        { year: 2026, month: 0 },
        { year: 2026, month: 3 },
        { year: 2026, month: 6 },
        { year: 2026, month: 9 },
        { year: 2027, month: 0 },
        { year: 2027, month: 6 }
    ];
}

function formatPlanFactMonthLabel(date) {
    return formatMonthName(date.getMonth());
}

function renderPlanFactProjectRow(project, startDate, endDate) {
    const rowClass = project.deviation.type === 'risk'
        ? 'planfact-row--risk'
        : project.status.type === 'done'
            ? 'planfact-row--done'
            : project.status.type === 'active'
                ? 'planfact-row--active'
                : 'planfact-row--empty';
    const projectNoteParts = [
        project.center ? `<span class="planfact-note-token">центр: ${escapeHtml(project.center)}</span>` : '',
        project.actualHours > 0 ? `<span class="planfact-note-token" data-fact-field>факт ${escapeHtml(Math.round(project.actualHours).toLocaleString('ru-RU'))} ч</span>` : '',
        project.planHours > 0 ? `<span class="planfact-note-token">план ${escapeHtml(Math.round(project.planHours).toLocaleString('ru-RU'))} ч</span>` : '',
        project.employeeCount > 0 ? `<span class="planfact-note-token" data-fact-field>${escapeHtml(project.employeeCount)} ${escapeHtml(formatEmployeeWord(project.employeeCount))}</span>` : ''
    ].filter(Boolean);
    const projectNoteHtml = projectNoteParts.length > 0
        ? projectNoteParts.join('')
        : `<span class="planfact-note-token" data-fact-field>статус: ${escapeHtml(project.status.label)}</span>`;

    return `
        <div class="planfact-row ${rowClass}">
            <div class="planfact-project-cell">
                <div class="planfact-project-name">${escapeHtml(project.id)} · ${escapeHtml(project.name)}</div>
                <div class="planfact-project-note">${projectNoteHtml}</div>
            </div>
            <div class="planfact-timeline-cell">
                <div class="planfact-timeline planfact-timeline--dual">
                    ${renderPlanFactTimelineLine('план', project.planStart, project.planEnd, project.planStartText, project.planEndText, startDate, endDate, 'plan')}
                    ${renderPlanFactTimelineLine('факт', project.factStart, project.factEnd, project.factStartText, project.factEndText, startDate, endDate, 'fact')}
                </div>
            </div>
            <div class="planfact-stages-cell" data-fact-field>
                ${project.stages.map(stage => `
                    <span class="planfact-stage planfact-stage--${stage.status.type}">
                        <i></i>
                        <span class="planfact-stage__label" title="${escapeHtml(stage.label)}">${escapeHtml(stage.shortLabel || stage.label)}</span>
                        <span class="planfact-stage__status">${escapeHtml(stage.status.label)}</span>
                    </span>
                `).join('')}
            </div>
            <div class="planfact-status-cell" data-label="Статус" data-fact-field>
                <span class="planfact-project-status planfact-project-status--${project.status.type}">${escapeHtml(project.status.label)}</span>
            </div>
            <div class="planfact-plan-cell" data-label="План">
                <span class="planfact-plan-badge planfact-plan-badge--${project.planFlag.type}">${escapeHtml(project.planFlag.label)}</span>
            </div>
            <div class="planfact-deviation-cell" data-label="Отклонение" data-fact-field>
                <span class="planfact-status planfact-status--${project.deviation.type}">${escapeHtml(project.deviation.label)}</span>
            </div>
        </div>
    `;
}

function renderPlanFactTimelineLine(label, start, end, startText, endText, rangeStart, rangeEnd, type) {
    const hasStart = start instanceof Date && !isNaN(start.getTime());
    const hasEnd = end instanceof Date && !isNaN(end.getTime());

    if (!hasStart && !hasEnd) {
        return `
            <div class="planfact-timeline-line planfact-timeline-line--${type}" ${type === 'fact' ? 'data-fact-field' : ''}>
                <span class="planfact-timeline-line__label">${label}</span>
                <div class="planfact-timeline-line__rail">
                    ${renderPlanFactYearBoundaries(rangeStart, rangeEnd)}
                    <span class="planfact-timeline__no-date">-</span>
                </div>
            </div>
        `;
    }

    const lineStart = hasStart ? start : end;
    const lineEnd = hasEnd ? end : start;
    const startPosition = getDatePositionPercent(lineStart, rangeStart, rangeEnd);
    const endPosition = getDatePositionPercent(lineEnd, rangeStart, rangeEnd);
    const left = Math.min(startPosition, endPosition);
    const width = Math.max(2, Math.abs(endPosition - startPosition));
    const markerPosition = hasEnd ? endPosition : startPosition;
    const markerText = hasEnd ? endText : startText;

    return `
        <div class="planfact-timeline-line planfact-timeline-line--${type}" ${type === 'fact' ? 'data-fact-field' : ''}>
            <span class="planfact-timeline-line__label">${label}</span>
            <div class="planfact-timeline-line__rail">
                ${renderPlanFactYearBoundaries(rangeStart, rangeEnd)}
                <div class="planfact-timeline__track"></div>
                <div class="planfact-timeline__progress planfact-timeline__progress--${type}" style="left: ${left}%; width: ${width}%;"></div>
                <div class="planfact-timeline__marker planfact-timeline__marker--${type}" style="left: ${markerPosition}%;"><span>${escapeHtml(markerText)}</span></div>
            </div>
        </div>
    `;
}

function setNavbarMenuOpen(isOpen) {
    const navbar = document.querySelector('.navbar');
    const toggle = document.getElementById('nav-menu-toggle');
    if (!navbar || !toggle) return;

    navbar.classList.toggle('navbar--menu-open', isOpen);
    toggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    toggle.setAttribute('aria-label', isOpen ? 'Закрыть меню разделов' : 'Открыть меню разделов');
}

function closeNavbarMenu() {
    setNavbarMenuOpen(false);
}

// Функция-заглушка: раньше обновляла скрытую секцию с путём к файлу.
// Сейчас не нужна, но оставляю как no-op, чтобы не ломать вызовы из loadSheetData.
function updateFileInfo() {}

// Toast-уведомления: показывают статус/ошибки в правом нижнем углу.
// Виды: loading (синий, не исчезает), success (зелёный, 2.5с), error (красный, 6с), warning (жёлтый, 4с).
let currentLoadingToast = null;

function showToast(message, type = 'info', duration = 3000) {
    const container = document.getElementById('toast-container');
    if (!container) return null;

    const toast = document.createElement('div');
    toast.className = `toast toast--${type}`;
    toast.setAttribute('role', type === 'error' ? 'alert' : 'status');

    const iconMap = {
        loading: '⟳',
        success: '✓',
        error: '✕',
        warning: '⚠',
        info: 'ℹ'
    };

    toast.innerHTML = `
        <span class="toast__icon">${iconMap[type] || iconMap.info}</span>
        <span class="toast__message"></span>
        <button type="button" class="toast__close" aria-label="Закрыть">✕</button>
    `;
    toast.querySelector('.toast__message').textContent = message;

    container.appendChild(toast);

    // Плавное появление на следующем кадре
    requestAnimationFrame(() => toast.classList.add('toast--visible'));

    const remove = () => {
        toast.classList.remove('toast--visible');
        setTimeout(() => toast.remove(), 250);
    };

    toast.querySelector('.toast__close').addEventListener('click', remove);

    if (duration > 0) {
        setTimeout(remove, duration);
    }

    return { el: toast, remove };
}

function setStatus(message, type) {
    // Закрываем предыдущий loading-toast, если он был
    if (currentLoadingToast) {
        currentLoadingToast.remove();
        currentLoadingToast = null;
    }

    if (type === 'loading') {
        currentLoadingToast = showToast(message, 'loading', 0); // без авто-закрытия
    } else if (type === 'error') {
        showToast(message, 'error', 6000);
    } else if (type === 'warning') {
        showToast(message, 'warning', 4000);
    } else if (type === 'success') {
        showToast(message, 'success', 2500);
    } else {
        showToast(message, 'info', 3000);
    }
}

// ============================================================
// ГЛОБАЛЬНЫЕ ФУНКЦИИ
// ============================================================

window.toggleEmployeesList = function(button) {
    const header = button.parentElement;
    const dataItem = header.parentElement;
    const list = dataItem.querySelector('.employees-list');
    const toggleBtn = header.querySelector('.toggle-btn');
    
    if (list && toggleBtn) {
        list.classList.toggle('show');
        toggleBtn.textContent = list.classList.contains('show') ? '▲' : '▼';
    }
};

window.toggleProjectDetails = function(element, status, type) {
    const allWrappers = document.querySelectorAll('.projects-details-wrapper');
    let targetWrapper = null;
    
    for (let wrapper of allWrappers) {
        if (wrapper.dataset.status === status && wrapper.dataset.type === type) {
            targetWrapper = wrapper;
            break;
        }
    }
    
    if (!targetWrapper) return;
    
    const isCurrentlyOpen = targetWrapper.style.display === 'block';
    
    allWrappers.forEach(wrapper => {
        wrapper.style.display = 'none';
    });
    
    document.querySelectorAll('.projects-count-cell').forEach(c => {
        c.classList.remove('expanded');
        const icon = c.querySelector('.toggle-icon-small');
        if (icon) icon.textContent = '▼';
    });
    
    if (!isCurrentlyOpen) {
        expandedProjectDetailsKey = getProjectDetailsExpandKey(status, type);
        targetWrapper.style.display = 'block';
        if (element) {
            element.classList.add('expanded');
            const icon = element.querySelector('.toggle-icon-small');
            if (icon) icon.textContent = '▲';
        }
    } else {
        expandedProjectDetailsKey = null;
    }
};

// ============================================================
// ИНИЦИАЛИЗАЦИЯ
// ============================================================

document.addEventListener('DOMContentLoaded', () => {
    log('DOM загружен, инициализация...');
    
    loadExcel();

    const navbar = document.querySelector('.navbar');
    const navMenuToggle = document.getElementById('nav-menu-toggle');
    if (navMenuToggle) {
        navMenuToggle.addEventListener('click', (e) => {
            e.stopPropagation();
            const isOpen = navMenuToggle.getAttribute('aria-expanded') === 'true';
            setNavbarMenuOpen(!isOpen);
        });
    }
    
    const refreshBtn = document.getElementById('refresh-btn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => {
            if (!refreshBtn.disabled) {
                loadExcel();
            }
        });
    }
    
    document.querySelectorAll('.nav-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const sheetName = e.currentTarget.dataset.sheet;
            if (sheetName && workbook) {
                loadSheetData(sheetName);
                closeNavbarMenu();
            } else if (!workbook) {
                setStatus('Сначала дождитесь загрузки файла', 'warning');
            }
        });
    });

    document.addEventListener('change', (e) => {
        if (e.target && e.target.matches('[data-pp-fact-toggle]')) {
            handleProductionProgramFactToggle(e.target);
            return;
        }

        if (e.target && e.target.classList.contains('fact-toggle__input')) {
            handleFactToggleChange(e.target);
            return;
        }

        if (e.target && e.target.matches('[data-pp-filter]')) {
            handleProductionProgramFilterChange(e.target);
        }
    });

    document.addEventListener('click', (e) => {
        if (e.target && e.target.matches('[data-pp-filter-reset]')) {
            resetProductionProgramFilters();
            return;
        }

        if (!navbar || !navbar.classList.contains('navbar--menu-open')) return;
        if (e.target.closest('.navbar')) return;
        closeNavbarMenu();
    });

    document.addEventListener('toggle', (e) => {
        if (!e.target || !e.target.classList || !e.target.classList.contains('planfact-details')) return;
        if (currentSheet !== 'ПП') return;

        productionProgramDetailsExpanded = e.target.open;
        renderProductionProgramPlanFact();
    }, true);

    // Горячие клавиши: ← → — переключают центры по порядку
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            closeNavbarMenu();
            return;
        }

        if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;

        // Не мешаем вводу в полях и когда открыт диалог/модалка
        const target = e.target;
        const isEditable = target && (
            target.tagName === 'INPUT' ||
            target.tagName === 'TEXTAREA' ||
            target.tagName === 'SELECT' ||
            target.isContentEditable
        );
        if (isEditable) return;
        const modalOpen = document.querySelector('.filter-modal--open');
        if (modalOpen) return;

        if (!workbook) return;

        const buttons = Array.from(document.querySelectorAll('.nav-btn'));
        if (buttons.length === 0) return;

        const currentIdx = buttons.findIndex(btn => btn.dataset.sheet === currentSheet);
        const startIdx = currentIdx === -1 ? 0 : currentIdx;
        const step = e.key === 'ArrowRight' ? 1 : -1;
        const nextIdx = (startIdx + step + buttons.length) % buttons.length;
        const nextSheet = buttons[nextIdx].dataset.sheet;

        if (nextSheet) {
            e.preventDefault();
            loadSheetData(nextSheet);
        }
    });
});

window.addEventListener('beforeunload', () => {
    chartManager.destroyAll();
});
