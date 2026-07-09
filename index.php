<?php
header('Content-Type: index.php; charset=UTF-8');
function asset($path) {
    $relativePath = ltrim($path, '/');
    $fullPath = __DIR__ . '/' . $relativePath;
    if (file_exists($fullPath)) {
        return $relativePath . '?v=' . filemtime($fullPath);
    }
    return $relativePath;
}
?>

<!DOCTYPE html>
<html lang="ru">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="<?= asset('style.css') ?>">
    <title>Отчет по центрам</title>
</head>

<body>
<div class="page-wrapper">
    <div class="container">
        <!-- Навигационная панель с кнопками переключения центров -->
        <div class="navbar">
            <div class="navbar__nav">
                <button id="nav-menu-toggle"
                        class="navbar__menu-toggle"
                        type="button"
                        aria-label="Открыть меню разделов"
                        aria-controls="navbar-buttons"
                        aria-expanded="false">
                    <span class="navbar__menu-icon" aria-hidden="true"></span>
                    <span id="nav-menu-current" class="navbar__menu-current">ТехБлок</span>
                </button>
                <div id="navbar-buttons" class="navbar__buttons">
                    <button class="nav-btn active" data-sheet="ТехБлок" title="Технический блок">ТехБлок</button>
                    <button class="nav-btn" data-sheet="ЦНГО" title="Центр нефтегазового оборудования">ЦНГО</button>
                    <button class="nav-btn" data-sheet="ЦДЦМ" title="Центр центробежных насосов, центробежных компрессоров и роторных машин">ЦДЦМ</button>
                    <button class="nav-btn" data-sheet="ЦДПМ" title="Центр деталей поршневых машин">ЦДПМ</button>
                    <button class="nav-btn" data-sheet="ЦОТ" title="Центр осевых турбомашин">ЦОТ</button>
                    <button class="nav-btn" data-sheet="УВП" title="Управление виртуального полигона">УВП</button>
                    <button class="nav-btn" data-sheet="ЛОИ" title="Управление лабораторными исследованиями">ЛОИ</button>
                    <button class="nav-btn" data-sheet="НКиТК" title="НКиТК">НК/ТК</button>
                    <button class="nav-btn" data-sheet="ПП" title="Производственная программа">ПП</button>
                </div>
            </div>
            <div class="navbar__meta">
                <div id="dashboard-updated-at" class="dashboard-updated-at" style="display: none;">
                    <span class="dashboard-updated-at__label">данные актуальны на</span>
                    <span id="dashboard-updated-at-time" class="dashboard-updated-at__time"></span>
                    <span id="dashboard-updated-at-value" class="dashboard-updated-at__value">—</span>
                </div>
                <button id="refresh-btn" class="navbar__refresh" type="button" title="Обновить данные" aria-label="Обновить данные">
                    <span class="navbar__refresh-icon">⟳</span>
                </button>
            </div>
        </div>

        <!-- Контейнер для всплывающих уведомлений -->
        <div id="toast-container" class="toast-container" aria-live="polite"></div>

        <!-- Основной контент для обычных страниц -->
        <div id="main-content">
            <header class="dashboard-header">
                <div class="dashboard-header__top">
                    <div class="dashboard-header__title-group">
                        <h1 class="center-title"><span class="skeleton skeleton--title"></span></h1>
                        <div id="header-employee-count" class="employee-stat" style="display: none;">
                            <span class="employee-stat__label">сотрудников</span>
                            <span id="total-employees-count" class="employee-stat__value">0</span>
                        </div>
                    </div>
                </div>
                <div class="header-info">
                    <h6><span class="skeleton skeleton--short"></span></h6>
                    <h6><span class="skeleton skeleton--short"></span></h6>
                </div>
            </header>

            <!-- Основная секция отчета -->
            <section class="report-section">
                <!-- Верхняя часть: Проекты, Табель и Загрузка -->
                <div class="upper-section">
                    <!-- Блок ТРУДОЗАТРАТЫ (посередине) -->
                    <div class="workload">
                        <!-- Блок ЗАГРУЗКА с диаграммами загрузки -->
                        <div class="block">
                            <div class="workload-header-with-count">
                                <div class="workload-header-title">
                                    <div class="workload-header-title-row">
                                        <h3>Загрузка центра</h3>
                                        <div id="workload-attention-badge" class="chart-title__status-badge workload-attention-badge" style="display: none;">
                                            Требует внимания • 0
                                        </div>
                                    </div>
                                    <span class="workload-title_comment">(без учета административных работ и простоя)</span>
                                </div>
                                <div class="workload-header-meta">
                                    <div id="workload-count-badge" class="workload-count-badge" style="display: none;">
                                        <span class="workload-count-icon">сотрудников без административного персонала: </span>
                                        <span id="workload-total-count" class="workload-count-number">0</span>
                                    </div>
                                </div>
                            </div>
                            <div class="charts-grid" id="charts-grid">
                                <!-- Карточки будут заполняться динамически через JavaScript -->
                            </div>
                            <div id="center-monthly-load-section" class="center-monthly-load-section"></div>
                            <div id="quarterly-trend-section" class="quarterly-trend-section"></div>
                        </div>

                        <!-- Блок ТАБЕЛЬ (справа) -->
                        <div class="report_card">
                            <!-- Блок ТАБЕЛЬ -->
                            <div class="block-employees">
                                <div class="block block--accent block--accent-timesheet" id="absences-block">
                                    <h3>Табель</h3>
                                    <div id="absences-container">
                                        <div class="skeleton-group">
                                            <span class="skeleton skeleton--line"></span>
                                            <span class="skeleton skeleton--short"></span>
                                            <span class="skeleton skeleton--line"></span>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <!-- Блок ЗАГРУЗКА ресурсов -->
                            <div class="blocks-container">
                                <div class="block block--accent block--accent-resources">
                                    <h3>Загрузка ресурсов</h3>
                                    <div id="loading-container">
                                        <div class="skeleton-group">
                                            <span class="skeleton skeleton--line"></span>
                                            <span class="skeleton skeleton--short"></span>
                                        </div>
                                    </div>
                                </div>
                            </div>
			    <!-- Блок УВЕДОМЛЕНИЕ -->
                            <section class="notification-section">
                                <div class="notification-block" id="notification-block" style="display: none;">
                                    <h3>Уведомление</h3>
                                    <div class="notification-content" id="notification-content">
                                        <!-- Текст уведомления будет загружен из Excel -->
                                    </div>
                                </div>
                                <div class="target-indicator-block" id="target-indicator-block" style="display: none;">
                                    <h3>Целевой показатель</h3>
                                    <div class="target-indicator-content" id="target-indicator-content">
                                        <!-- Целевой показатель будет рассчитан из Проекты_ТехБлок -->
                                    </div>
                                </div>
                            </section>
                        </div>
                    </div>
                    <!-- Блок ПРОЕКТЫ (слева) -->
                    <div class="block-projects">
                        <!-- Блок ПРОЕКТЫ -->
                        <div class="block" id="projects-block">
                            <h3>Проекты</h3>
                            <!-- Блок с легендой индикаторов -->
                            <div class="indicators-legend" id="indicators-legend" style="display: none;">
                                <div class="indicators-legend__content">
                                    <div class="legend-header-container">
                                        <h4 class="legend-main-title">Легенда индикаторов</h4>
                                    </div>
                                    <div class="legend-items" id="legend-items">
                                        <!-- Содержимое будет добавлено через JavaScript -->
                                    </div>
                                </div>
                                <div class="projects-target-filter-slot" id="projects-target-filter-slot"></div>
                            </div>
                            <div id="projects-container">
                                <div class="skeleton-group">
                                    <span class="skeleton skeleton--title"></span>
                                    <span class="skeleton skeleton--line"></span>
                                    <span class="skeleton skeleton--line"></span>
                                    <span class="skeleton skeleton--short"></span>
                                    <span class="skeleton skeleton--line"></span>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </section>
        </div>

        <!-- Контейнер для страницы производственной программы -->
        <div id="production-program-content" style="display: none;">
            <header class="dashboard-header production-program-header">
                <div class="dashboard-header__top">
                    <div class="dashboard-header__title-group">
                        <h1 class="center-title">Производственная программа</h1>
                    </div>
                </div>
            </header>
            <section class="report-section">
                <div class="block production-program-block" id="production-program-projects-block">
                    <h3>План-факт по проектам</h3>
                    <div id="production-program-projects-container">
                        <div class="skeleton-group">
                            <span class="skeleton skeleton--title"></span>
                            <span class="skeleton skeleton--line"></span>
                            <span class="skeleton skeleton--line"></span>
                        </div>
                    </div>
                </div>
            </section>
        </div>

        <!-- Контейнер для страницы НКиТК (изначально скрыт) -->
        <div id="nkitk-content" style="display: none;">
            <header>
                <h1 class="center-title">НКиТК - Данные</h1>
            </header>
            <section class="report-section">
                <div class="block">
                    <!--<h3>Таблица данных (столбцы A:H)</h3>-->
                    <div class="table-wrapper" style="overflow-x: auto; margin-top: 20px;">
                        <div id="nkitk-table-container">
                            <p>Загрузка данных...</p>
                        </div>
                    </div>
                </div>
            </section>
        </div>

        <!-- Контейнер для страницы самопроверки расчетов -->
        <div id="calculations-content" style="display: none;">
            <header>
                <h1 class="center-title">Расчеты и источники данных</h1>
            </header>
            <section class="report-section">
                <div id="calculations-container">
                    <div class="skeleton-group">
                        <span class="skeleton skeleton--title"></span>
                        <span class="skeleton skeleton--line"></span>
                        <span class="skeleton skeleton--line"></span>
                    </div>
                </div>
            </section>
        </div>
    </div>

    <!-- Подключаем внешние библиотеки (актуальные версии) -->
    <script defer src="https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>

    <!-- Подключаем наш внешний JavaScript файл -->
    <script defer src="<?= asset('script.js') ?>"></script>
</body>
</html>
