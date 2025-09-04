document.addEventListener("DOMContentLoaded", () => {
    const fileUpload = document.getElementById('file-upload');
    const fileNameSpan = document.getElementById('file-name');
    const loader = document.getElementById('loader');
    const customerChartContainer = document.getElementById('customer-chart-container');
    const partDetailChartContainer = document.getElementById('part-detail-chart-container');
    const partDetailTitleText = document.getElementById('part-detail-title-text');
    const messageArea = document.getElementById('message-area');
    const messageText = document.getElementById('message-text');
    const truncationWarning = document.getElementById('truncation-warning');
    const startDateFilter = document.getElementById('start-date-filter');
    const endDateFilter = document.getElementById('end-date-filter');
    const filterButton = document.getElementById('filter-button');
    const clearFilterButton = document.getElementById('clear-filter-button');
    const statusFilterContainer = document.getElementById('status-filters');
    const sortOptions = document.getElementById('sort-options');
    const customerSearchInput = document.getElementById('customer-search');
    const customerFilterGantt = document.getElementById('customer-filter-gantt');
    const progressFilterGantt = document.getElementById('progress-filter-gantt');
    const searchSn = document.getElementById('search-sn');
    const searchPrDraft = document.getElementById('search-pr-draft');

    const tabsContainer = document.getElementById('tabs-container');
    const ganttTabButton = document.getElementById('gantt-tab-button');
    const tableTabButton = document.getElementById('table-tab-button');
    const alarmTabButton = document.getElementById('alarm-tab-button');
    const kpiTabButton = document.getElementById('kpi-tab-button');
    const forecastTabButton = document.getElementById('forecast-tab-button');
    const completionTabButton = document.getElementById('completion-tab-button');
    const ganttDetailTabButton = document.getElementById('gantt-detail-tab-button');
    const monthlyForecastTabButton = document.getElementById('monthly-forecast-tab-button');

    const ganttTabContent = document.getElementById('gantt-tab-content');
    const tableTabContent = document.getElementById('table-tab-content');
    const alarmTabContent = document.getElementById('alarm-tab-content');
    const kpiTabContent = document.getElementById('kpi-tab-content');
    const forecastTabContent = document.getElementById('forecast-tab-content');
    const completionTabContent = document.getElementById('completion-tab-content');
    const ganttDetailTabContent = document.getElementById('gantt-detail-tab-content');
    const monthlyForecastTabContent = document.getElementById('monthly-forecast-tab-content');

    const progressTableBody = document.getElementById('progress-table-body');
    const kpiTableBody = document.getElementById('kpi-table-body');
    const forecastTableBody = document.getElementById('forecast-table-body');
    const completionTableBody = document.getElementById('completion-table-body');

    const progressTableControls = document.getElementById('progress-table-controls');
    const customerFilterProgress = document.getElementById('customer-filter-progress');
    const progressFilter = document.getElementById('progress-filter');
    const progressSort = document.getElementById('progress-sort');

    const kpiTableControls = document.getElementById('kpi-table-controls');
    const kpiFilterCareOf = document.getElementById('kpi-filter-care-of');
    const kpiFilterProgress = document.getElementById('kpi-filter-progress');

    const forecastTableControls = document.getElementById('forecast-table-controls');
    const forecastMonthFilter = document.getElementById('forecast-month-filter');

    const completionTableControls = document.getElementById('completion-table-controls');
    const completionMonthFilter = document.getElementById('completion-month-filter');
    const customerFilterCompletion = document.getElementById('customer-filter-completion');
    const completionSort = document.getElementById('completion-sort');
    const completionPercentageFilter = document.getElementById('completion-percentage-filter');

    // Gantt Detail Tab elements
    const customerFilterGanttDetail = document.getElementById('customer-filter-gantt-detail');
    const ganttDetailMainContainer = document.getElementById('gantt-detail-main-container');
    const searchSnDetail = document.getElementById('search-sn-detail');
    const searchProgressDetail = document.getElementById('search-progress-detail');
    const ganttDetailTitleText = document.getElementById('gantt-detail-title-text');
    const ganttDetailTitleStats = document.getElementById('gantt-detail-title-stats');

    // Monthly Forecast elements
    const monthlyForecastChartContainer = document.getElementById('monthly-forecast-chart-container');
    const forecastChartYearFilter = document.getElementById('forecast-chart-year-filter');
    const forecastChartMonthFilter = document.getElementById('forecast-chart-month-filter');
    const forecastChartApplyBtn = document.getElementById('forecast-chart-apply-filter');
    const monthlyForecastChartMessage = document.getElementById('monthly-forecast-chart-message');
    const monthlyForecastLegend = document.getElementById('monthly-forecast-legend');
    const monthlyForecastTitleText = document.getElementById('monthly-forecast-title-text');
    const miniCompletionTableContainer = document.getElementById('mini-completion-table-container');


    // Print Buttons
    const printGanttCustomerBtn = document.getElementById('print-gantt-customer-btn');
    const printGanttPartBtn = document.getElementById('print-gantt-part-btn');
    const printGanttDetailBtn = document.getElementById('print-gantt-detail-btn');
    const printMonthlyForecastBtn = document.getElementById('print-monthly-forecast-btn');
    const printProgressBtn = document.getElementById('print-progress-btn');
    const printAlarmBtn = document.getElementById('print-alarm-btn');
    const printKpiBtn = document.getElementById('print-kpi-btn');
    const printForecastBtn = document.getElementById('print-forecast-btn');
    const printCompletionBtn = document.getElementById('print-completion-btn');

    const tooltip = d3.select("body").append("div").attr("class", "tooltip");
    const DISPLAY_LIMIT = 200;

    let fullCustomerData = [];
    let fullRawData = [];
    let currentDetailData = [];
    let excelHeaderMap = {};
    let currentStatusFilter = 'delay';
    const progressStatuses = ['Cancelled', 'Completed', 'N/A', 'Shipped', 'Use Local', 'Use Stock', 'Wait Approve', 'Wait PO', 'Wait POS', 'Wait PR', 'Wait QO', 'WIP'];

    /**
     * A robust, central date parsing function.
     * It prioritizes MMM-YY and DD-MM-YYYY formats and handles other common cases like Excel serial numbers.
     * @param {*} val The value from the Excel cell.
     * @returns {Date|null} A valid Date object or null if parsing fails.
     */
    function parseExcelDate(val) {
        // 1. If it's already a valid Date object, return it.
        if (val instanceof Date && !isNaN(val)) {
            return val;
        }

        // 2. If it's an Excel serial number, convert it.
        if (typeof val === 'number') {
            // This formula converts Excel's serial date number to a JS Date.
            return new Date(Date.UTC(0, 0, val - 1));
        }

        // 3. If it's a string, attempt parsing.
        if (typeof val === 'string' && val.trim() !== '') {
            const str = val.trim();

            // A. Prioritize MMM-YY format (e.g., 'Sep-25')
            const mmmYyPattern = /^([a-zA-Z]{3})[-/](\d{2})$/;
            let match = str.match(mmmYyPattern);
            if (match) {
                const monthStr = match[1];
                const yearShort = parseInt(match[2], 10);
                const year = 2000 + yearShort; // Assumes 21st century

                const monthMap = { jan: 0, feb: 1, mar: 2, apr: 3, may: 4, jun: 5, jul: 6, aug: 7, sep: 8, oct: 9, nov: 10, dec: 11 };
                const monthIndex = monthMap[monthStr.toLowerCase()];

                if (monthIndex !== undefined) {
                    return new Date(Date.UTC(year, monthIndex, 1));
                }
            }

            // B. Prioritize DD-MM-YYYY format (e.g., '21-09-2025')
            const dmyPattern = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/;
            match = str.match(dmyPattern);
            if (match) {
                const day = parseInt(match[1], 10);
                const month = parseInt(match[2], 10); // Month is 1-indexed
                const year = parseInt(match[3], 10);

                if (day > 0 && day <= 31 && month > 0 && month <= 12) {
                    const date = new Date(Date.UTC(year, month - 1, day));
                    // Final check to ensure parts are valid (e.g., no 31st of Feb)
                    if (date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day) {
                        return date;
                    }
                }
            }

            // C. Fallback to native parser for other formats (like YYYY-MM-DD)
            const parsed = new Date(str);
            if (!isNaN(parsed)) {
                return parsed;
            }
        }

        // 4. If all else fails, return null.
        return null;
    }


    fileUpload.addEventListener('change', handleFile);

    filterButton.addEventListener('click', applyAllFilters);
    clearFilterButton.addEventListener('click', clearFilters);
    sortOptions.addEventListener('change', applyAllFilters);
    customerSearchInput.addEventListener('input', applyAllFilters);
    customerFilterGantt.addEventListener('change', applyAllFilters);
    progressFilterGantt.addEventListener('change', applyAllFilters);
    searchSn.addEventListener('input', applyAllFilters);
    searchPrDraft.addEventListener('input', applyAllFilters);

    statusFilterContainer.addEventListener('click', (e) => {
        if (e.target.matches('.status-filter-btn')) {
            statusFilterContainer.querySelector('.active-filter')?.classList.remove('active-filter');
            e.target.classList.add('active-filter');
            currentStatusFilter = e.target.dataset.status;
            applyAllFilters();
        }
    });

    function applyAllFilters() {
        applyFilters();
        renderProgressTable();
        renderKpiReport();
        renderForecastTable();
        renderCompletionTable();
    }

    function switchTab(activeButton, activeContent) {
         [ganttTabButton, tableTabButton, alarmTabButton, kpiTabButton, forecastTabButton, completionTabButton, ganttDetailTabButton, monthlyForecastTabButton].forEach(btn => btn.classList.remove('active'));
         [ganttTabContent, tableTabContent, alarmTabContent, kpiTabContent, forecastTabContent, completionTabContent, ganttDetailTabContent, monthlyForecastTabContent].forEach(content => content.classList.remove('active'));
         activeButton.classList.add('active');
         activeContent.classList.add('active');
    }
    ganttTabButton.addEventListener('click', () => switchTab(ganttTabButton, ganttTabContent));
    ganttDetailTabButton.addEventListener('click', () => switchTab(ganttDetailTabButton, ganttDetailTabContent));
    monthlyForecastTabButton.addEventListener('click', () => switchTab(monthlyForecastTabButton, monthlyForecastTabContent));
    tableTabButton.addEventListener('click', () => switchTab(tableTabButton, tableTabContent));
    kpiTabButton.addEventListener('click', () => switchTab(kpiTabButton, kpiTabContent));
    forecastTabButton.addEventListener('click', () => switchTab(forecastTabButton, forecastTabContent));
    completionTabButton.addEventListener('click', () => switchTab(completionTabButton, completionTabContent));

    alarmTabButton.addEventListener('click', () => {
        switchTab(alarmTabButton, alarmTabContent);
        const isRendered = document.querySelector("#arrived-chart").hasChildNodes();
        if (!isRendered && fullRawData.length > 0) {
            renderArrivalAlarms();
        }
    });

    customerFilterProgress.addEventListener('change', renderProgressTable);
    progressFilter.addEventListener('change', renderProgressTable);
    progressSort.addEventListener('change', renderProgressTable);

    kpiFilterCareOf.addEventListener('change', renderKpiReport);
    kpiFilterProgress.addEventListener('change', renderKpiReport);
    forecastMonthFilter.addEventListener('change', renderForecastTable);

    completionMonthFilter.addEventListener('change', renderCompletionTable);
    customerFilterCompletion.addEventListener('change', renderCompletionTable);
    completionSort.addEventListener('change', renderCompletionTable);
    completionPercentageFilter.addEventListener('change', renderCompletionTable);


    // Event listeners for Gantt Detail Tab
    customerFilterGanttDetail.addEventListener('change', renderFullDetailGantt);
    searchSnDetail.addEventListener('input', renderFullDetailGantt);
    searchProgressDetail.addEventListener('input', renderFullDetailGantt);

    // Event listener for Monthly Forecast Tab
    forecastChartApplyBtn.addEventListener('click', renderMonthlyForecastChart);

    // Event Listeners for all Print Buttons
    printGanttCustomerBtn.addEventListener('click', () => printElementAsPng('customer-chart-container', 'Gantt_Customer_View.png'));
    printGanttPartBtn.addEventListener('click', () => printElementAsPng('part-detail-chart-container', 'Gantt_Part_Detail_View.png'));
    printGanttDetailBtn.addEventListener('click', () => {
        const customerName = customerFilterGanttDetail.value || "Report";
        const fileName = `Gantt_Details_${customerName.replace(/ /g, '_')}.png`;
        printElementAsPng('gantt-detail-main-container', fileName);
    });
    printMonthlyForecastBtn.addEventListener('click', () => {
        const year = forecastChartYearFilter.value;
        const month = forecastChartMonthFilter.options[forecastChartMonthFilter.selectedIndex].text;
        const fileName = `Monthly_Forecast_${month}_${year}.png`;
        printElementAsPng('monthly-forecast-chart-container', fileName);
    });
    printProgressBtn.addEventListener('click', () => printElementAsPng('progress-table-container', 'Progress_Report.png'));
    printAlarmBtn.addEventListener('click', () => printElementAsPng('alarm-tab-content', 'Arrival_Alarm_Report.png'));
    printKpiBtn.addEventListener('click', () => printElementAsPng('kpi-table-container', 'KPI_Delay_Report.png'));
    printForecastBtn.addEventListener('click', () => printElementAsPng('forecast-table-container', 'Forecast_Report.png'));
    printCompletionBtn.addEventListener('click', () => printElementAsPng('completion-table-container', 'Completion_Report.png'));


    function handleFile(e) {
        const file = e.target.files[0];
        if (!file) return;
        fileNameSpan.textContent = file.name;
        loader.classList.remove('hidden');
        [messageArea, customerChartContainer, truncationWarning, partDetailChartContainer, tabsContainer, progressTableControls, kpiTableControls, forecastTableControls, completionTableControls, ganttDetailMainContainer, monthlyForecastChartContainer].forEach(el => el.classList.add('hidden'));
        d3.select("#gantt-chart").selectAll("*").remove();
        d3.select("#gantt-detail-chart").selectAll("*").remove();
        d3.select("#monthly-forecast-chart").selectAll("*").remove();
        miniCompletionTableContainer.classList.add('hidden');
        miniCompletionTableContainer.innerHTML = '';
        progressTableBody.innerHTML = '';
        kpiTableBody.innerHTML = '';
        forecastTableBody.innerHTML = '';
        completionTableBody.innerHTML = '';
        d3.select("#arrived-chart").selectAll("*").remove();
        d3.select("#arriving-1-week-chart").selectAll("*").remove();
        d3.select("#arriving-2-weeks-chart").selectAll("*").remove();

        const workerScript = document.getElementById('data-worker').textContent;
        const workerBlob = new Blob([workerScript], { type: 'application/javascript' });
        const workerUrl = URL.createObjectURL(workerBlob);
        const worker = new Worker(workerUrl);

        worker.onmessage = (event) => {
            const { customerData, rawData, headerMap, error } = event.data;
            loader.classList.add('hidden');
            if (error) {
                showError(error);
                return;
            }
            tabsContainer.classList.remove('hidden');
            progressTableControls.classList.remove('hidden');
            kpiTableControls.classList.remove('hidden');
            forecastTableControls.classList.remove('hidden');
            completionTableControls.classList.remove('hidden');
            ganttDetailMainContainer.classList.remove('hidden');
            monthlyForecastChartContainer.classList.remove('hidden');

            excelHeaderMap = headerMap;
            
            // Pre-process the raw data to standardize all dates upon loading
            const dateColumns = ['ETA', 'PO', 'OD', 'ETF', 'ETD', 'Forecast', 'Deliver Date', 'Sale Order No.'];
            rawData.forEach(row => {
                dateColumns.forEach(colName => {
                    const header = excelHeaderMap[colName];
                    if (header && row[header] !== undefined) {
                        // Overwrite the original value with a proper Date object or null
                        row[header] = parseExcelDate(row[header]);
                    }
                });
            });
            fullRawData = rawData;

            // This data for the main gantt chart also needs its dates standardized
            fullCustomerData = customerData.map(d => ({...d, 
                saleOrderDate: parseExcelDate(d.saleOrderDate), 
                deliverDate: parseExcelDate(d.deliverDate), 
                etaDate: parseExcelDate(d.etaDate)
            }));


            progressFilter.innerHTML = '<option value="all">All Statuses</option>';
            progressFilterGantt.innerHTML = '<option value="all">All Statuses</option>';
            kpiFilterProgress.innerHTML = '<option value="all">All Statuses</option>';
            progressStatuses.forEach(status => {
                const option = document.createElement('option');
                option.value = status;
                option.textContent = status;
                progressFilter.appendChild(option.cloneNode(true));
                progressFilterGantt.appendChild(option.cloneNode(true));
                kpiFilterProgress.appendChild(option.cloneNode(true));
            });

            const customerNames = [...new Set(fullRawData.map(row => row[excelHeaderMap['Customer Name']]))].sort();
            customerFilterProgress.innerHTML = '<option value="all">All Customers</option>';
            customerFilterGantt.innerHTML = '<option value="all">All Customers</option>';
            customerFilterGanttDetail.innerHTML = '<option value="">-- Select a Customer --</option>';
            customerFilterCompletion.innerHTML = '';
            customerNames.forEach(name => {
                if(name) {
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    customerFilterProgress.appendChild(option.cloneNode(true));
                    customerFilterGantt.appendChild(option.cloneNode(true));
                    customerFilterGanttDetail.appendChild(option.cloneNode(true));
                    customerFilterCompletion.appendChild(option);
                }
            });

            const careOfNames = [...new Set(fullRawData.map(row => (row[excelHeaderMap['Care Of']] || "")).filter(name => name.trim() !== ''))].sort();
            kpiFilterCareOf.innerHTML = '<option value="all">All Staff</option>';
            careOfNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                kpiFilterCareOf.appendChild(option);
            });

            const forecastMonths = new Set();
            const forecastYears = new Set();
            fullRawData.forEach(row => {
                // Logic is now simpler as dates are pre-parsed
                const forecastDate = row[excelHeaderMap['Forecast']];
                if (forecastDate) {
                    // CORRECTED: Format the date using UTC to avoid timezone shifts
                    const formattedMonth = forecastDate.toLocaleString('default', { month: 'short', year: 'numeric', timeZone: 'UTC' });
                    forecastMonths.add(formattedMonth);
                    forecastYears.add(forecastDate.getUTCFullYear());
                }
            });

            forecastMonthFilter.innerHTML = '<option value="all">All Months</option>';
            completionMonthFilter.innerHTML = '<option value="all">All Months</option>';
            const sortedMonths = Array.from(forecastMonths).sort((a,b) => new Date(`1 ${a}`) - new Date(`1 ${b}`));
            sortedMonths.forEach(month => {
                const option = document.createElement('option');
                option.value = month;
                option.textContent = month;
                forecastMonthFilter.appendChild(option.cloneNode(true));
                completionMonthFilter.appendChild(option);
            });

            forecastChartYearFilter.innerHTML = '<option value="">-- Year --</option>';
            const sortedYears = Array.from(forecastYears).sort((a,b) => a - b);
            sortedYears.forEach(year => {
                 const option = document.createElement('option');
                 option.value = year;
                 option.textContent = year;
                 forecastChartYearFilter.appendChild(option);
            });

            applyAllFilters();

            URL.revokeObjectURL(workerUrl);
        };
        const reader = new FileReader();
        reader.onload = (event) => {
            worker.postMessage({ fileData: event.target.result }, [event.target.result]);
        };
        reader.readAsArrayBuffer(file);
    }

    function getUniversallyFilteredRawData() {
        let data = [...fullRawData];

        const searchTerm = customerSearchInput.value.toLowerCase();
        if (searchTerm) {
            data = data.filter(row => (row[excelHeaderMap['Customer Name']] || '').toLowerCase().includes(searchTerm));
        }

        const customerFilterValue = customerFilterGantt.value;
        if (customerFilterValue !== 'all') {
            data = data.filter(row => (row[excelHeaderMap['Customer Name']] || '') === customerFilterValue);
        }

        const progressFilterValue = progressFilterGantt.value;
        if (progressFilterValue !== 'all') {
            data = data.filter(row => (row[excelHeaderMap['Progress']] || '') === progressFilterValue);
        }

        const snSearchTerm = searchSn.value.toLowerCase();
        if (snSearchTerm) {
            data = data.filter(row => String(row[excelHeaderMap['S/N']] || '').toLowerCase().includes(snSearchTerm));
        }

        const prDraftSearchTerm = searchPrDraft.value.toLowerCase();
        if (prDraftSearchTerm && excelHeaderMap['PR Draft']) {
             data = data.filter(row => String(row[excelHeaderMap['PR Draft']] || '').toLowerCase().includes(prDraftSearchTerm));
        }
        return data;
    }

    function applyFilters() {
        const universallyFilteredRaw = getUniversallyFilteredRawData();
        const allowedSNs = new Set(universallyFilteredRaw.map(row => String(row[excelHeaderMap['S/N']])));

        let customerData = fullCustomerData.filter(d => allowedSNs.has(d.sn));

        partDetailChartContainer.classList.add('hidden');
        miniCompletionTableContainer.classList.add('hidden');
        miniCompletionTableContainer.innerHTML = '';

        if (currentStatusFilter !== 'all') {
            customerData = customerData.filter(d => d.status === currentStatusFilter);
        }

        const startDateString = startDateFilter.value;
        const endDateString = endDateFilter.value;
        if (startDateString && endDateString) {
            const startDate = new Date(startDateString);
            const endDate = new Date(endDateString);
            customerData = customerData.filter(d => {
                if (!d.saleOrderDate || !d.deliverDate) return false;
                const taskEndDate = d.isDelayed && d.etaDate ? d.etaDate : d.deliverDate;
                return d.saleOrderDate <= endDate && taskEndDate >= startDate;
            });
        }

        const sortBy = sortOptions.value;
        if (sortBy === 'name') {
            customerData.sort((a, b) => a.customer.localeCompare(b.customer));
        } else if (sortBy === 'max-delay') {
            customerData.sort((a, b) => b.maxDelay - a.maxDelay);
        }

        messageArea.classList.add('hidden');
        truncationWarning.classList.add('hidden');
        if (customerData.length > DISPLAY_LIMIT) {
            truncationWarning.textContent = `Displaying the first ${DISPLAY_LIMIT} customer rows for performance. Please use filters to narrow your search.`;
            truncationWarning.classList.remove('hidden');
            customerData = customerData.slice(0, DISPLAY_LIMIT);
        }

        if (customerData.length > 0) {
            customerChartContainer.classList.remove('hidden');
            requestAnimationFrame(() => createGanttChartForCustomer(customerData));
        } else {
            customerChartContainer.classList.add('hidden');
            if(fullCustomerData.length > 0) showError("No data matches the selected filters for the Gantt Chart.");
        }
    }

    function clearFilters() {
        startDateFilter.value = '';
        endDateFilter.value = '';
        customerSearchInput.value = '';
        customerFilterGantt.value = 'all';
        progressFilterGantt.value = 'all';
        searchSn.value = '';
        searchPrDraft.value = '';
        sortOptions.value = 'name';
        forecastMonthFilter.value = 'all';
        completionMonthFilter.value = 'all';
        completionPercentageFilter.value = 'all';
        completionSort.value = 'customer';
         Array.from(customerFilterCompletion.options).forEach(option => option.selected = false);
        statusFilterContainer.querySelector('.active-filter')?.classList.remove('active-filter');
        const delayButton = statusFilterContainer.querySelector('[data-status="delay"]');
        if (delayButton) {
           delayButton.classList.add('active-filter');
           currentStatusFilter = 'delay';
        }
        applyAllFilters();
    }

    function drawSharedElements(svg, xScale, margin, chartHeight, chartWidth) {
         const headerHeight = 120;
         svg.append("rect").attr("x", 0).attr("y", -headerHeight).attr("width", chartWidth).attr("height", headerHeight).attr("fill", "#f3f4f6");
         const [startDate, endDate] = xScale.domain();
         const durationInDays = (endDate - startDate) / (1000 * 60 * 60 * 24);
         let majorTickInterval, majorTickFormat, minorTickInterval, minorTickFormat, gridTickInterval;

         if (durationInDays > 70) { 
             majorTickInterval = d3.utcYear.every(1); majorTickFormat = d3.utcFormat("%Y");
             minorTickInterval = d3.utcMonth.every(1); minorTickFormat = d3.utcFormat("%b");
             gridTickInterval = d3.utcMonth.every(1);
         } else { 
             majorTickInterval = d3.utcWeek.every(1); majorTickFormat = d3.utcFormat("%b %d");
             minorTickInterval = d3.utcDay.every(1); minorTickFormat = d3.utcFormat("%d");
             gridTickInterval = d3.utcWeek.every(1);
         }

        const weekTickInterval = d3.utcWeek.every(1);
        const weekTicks = xScale.ticks(weekTickInterval);
        const weekGroup = svg.append("g").attr("class", "week-header-axis").attr("transform", `translate(0, ${-headerHeight})`);
        weekGroup.selectAll("g").data(weekTicks).enter().append("g").attr("class", "week-header-tick").each(function(d) {
            const tick = d3.select(this);
            const nextTickDate = weekTickInterval.offset(d, 1);
            const tickWidth = xScale(nextTickDate) - xScale(d);
            if (tickWidth > 20) {
                tick.append("rect").attr("x", xScale(d)).attr("y", 0).attr("width", tickWidth).attr("height", 40);
                const textX = xScale(d) + tickWidth / 2;
                const textY = 20;
                tick.append("text")
                    .attr("text-anchor", "middle")
                    .attr("transform", `translate(${textX}, ${textY}) rotate(-90)`)
                    .text(`W${d3.utcFormat("%U")(d)}`);
            }
        });
         
         const majorTicks = xScale.ticks(majorTickInterval);
         const majorGroup = svg.append("g").attr("class", "date-header-axis").attr("transform", `translate(0, ${-headerHeight + 40})`);
         majorGroup.selectAll("g").data(majorTicks).enter().append("g").attr("class", "date-header-tick").each(function(d) {
             const tick = d3.select(this), nextTickDate = majorTickInterval.offset(d, 1), tickWidth = xScale(nextTickDate) - xScale(d);
             if (tickWidth > 0) {
                 tick.append("rect").attr("x", xScale(d)).attr("y", 0).attr("width", tickWidth).attr("height", 40);
                 tick.append("text").attr("x", xScale(d) + tickWidth / 2).attr("y", 25).attr("text-anchor", "middle").text(majorTickFormat(d));
             }
         });
         const minorTicks = xScale.ticks(minorTickInterval);
         const minorGroup = svg.append("g").attr("class", "date-header-axis").attr("transform", `translate(0, ${-headerHeight + 80})`);
         minorGroup.selectAll("g").data(minorTicks).enter().append("g").attr("class", "date-header-tick").each(function(d) {
             const tick = d3.select(this), nextTickDate = minorTickInterval.offset(d, 1), tickWidth = xScale(nextTickDate) - xScale(d);
              if (tickWidth > 0 && tickWidth > 15) {
                 tick.append("rect").attr("x", xScale(d)).attr("y", 0).attr("width", tickWidth).attr("height", 40);
                 tick.append("text").attr("x", xScale(d) + tickWidth / 2).attr("y", 25).attr("text-anchor", "middle").text(minorTickFormat(d));
              }
         });
         svg.append("g").attr("class", "x-axis-grid").call(d3.axisTop(xScale).ticks(gridTickInterval).tickFormat("").tickSize(-chartHeight));
         const today = new Date();
         const todayUTC = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));

         if (todayUTC >= startDate && todayUTC <= endDate) {
             const todayX = xScale(todayUTC);
             svg.append("line").attr("x1", todayX).attr("x2", todayX).attr("y1", -headerHeight).attr("y2", chartHeight).attr("stroke", "darkblue").attr("stroke-width", 2).attr("stroke-dasharray", "6,3");
             svg.append("text").attr("x", todayX).attr("y", -headerHeight + 15).attr("text-anchor", "middle").attr("fill", "darkblue").style("font-size", "10px").style("font-weight", "bold").text("TODAY");
         }
    }

    function createGanttChartForCustomer(data) {
        const chartId = "#gantt-chart";
        d3.select(chartId).selectAll("*").remove();
        const margin = { top: 122, right: 40, bottom: 40, left: 350 };
        const containerWidth = document.querySelector(chartId).clientWidth;
        if (containerWidth <= 0) return;
        const chartHeight = Math.max(300, data.length * 60);
        const height = chartHeight + margin.top + margin.bottom;
        const width = containerWidth - margin.left - margin.right;
        const svgContainer = d3.select(chartId).append("svg").attr("width", containerWidth).attr("height", height);
        const svg = svgContainer.append("g").attr("transform", `translate(${margin.left},${margin.top})`);
        const plottableData = data.filter(d => d.saleOrderDate && d.deliverDate);
        if (plottableData.length === 0) return;
        const minDate = d3.min(plottableData, d => d.saleOrderDate);
        const maxDate = d3.max(plottableData, d => d.isDelayed && d.etaDate ? d.etaDate : d.deliverDate);
        const xScale = d3.scaleUtc().domain([minDate, maxDate]).range([0, width]).nice();
        const yScale = d3.scaleBand().domain(data.map(d => d.taskId)).range([0, chartHeight]).padding(0.4);
        drawSharedElements(svg, xScale, margin, chartHeight, width);
        
        const utcFormat = d3.utcFormat("%Y-%m-%d");

        const yHeader = svgContainer.append("g").attr("transform", `translate(0, 0)`);
        yHeader.append("rect").attr("x", 0).attr("y", 0).attr("width", margin.left).attr("height", margin.top -2).attr("fill", "#fef9c3").attr("stroke", "#9ca3af").attr("stroke-width", "1px");
        const yHeaderPositions = { no: 5, customer: (margin.left * 0.1) + 5, mcType: (margin.left * 0.45) + 5, sn: (margin.left * 0.70) + 5 };
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.no).attr("y", (margin.top -2) / 2 + 5).text("No.");
        yHeader.append("line").attr("x1", yHeaderPositions.customer - 5).attr("x2", yHeaderPositions.customer - 5).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af").attr("stroke-width", "1px");
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.customer).attr("y", (margin.top-2) / 2 + 5).text("Customer Name");
        yHeader.append("line").attr("x1", yHeaderPositions.mcType - 5).attr("x2", yHeaderPositions.mcType - 5).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af").attr("stroke-width", "1px");
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.mcType).attr("y", (margin.top-2) / 2 + 5).text("M/C Type");
        yHeader.append("line").attr("x1", yHeaderPositions.sn - 5).attr("x2", yHeaderPositions.sn - 5).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af").attr("stroke-width", "1px");
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.sn).attr("y", (margin.top-2) / 2 + 5).text("S/N");
        svgContainer.append("g").selectAll(".y-axis-label-group").data(data).enter().append("foreignObject")
            .attr("x", 0).attr("y", d => yScale(d.taskId) + margin.top)
            .attr("width", margin.left).attr("height", yScale.bandwidth())
            .append("xhtml:div").attr("class", "y-axis-label y-axis-row")
            .html((d, i) => `<div class="row-number" style="width:10%">${i + 1}</div><div class="customer-name" title="${d.customer}">${d.customer}</div><div class="mc-type" title="${d.mcType}">${d.mcType}</div><div class="sn" title="${d.sn}">${d.sn}</div>`);
        const bars = svg.selectAll(".bar-group").data(data).enter().append("g").attr("class", "bar-group").attr("transform", d => `translate(0, ${yScale(d.taskId)})`);
        function handleBarClick(event, d) {
            currentDetailData = d.detailRows;
            partDetailTitleText.textContent = `Part Details for ${d.customer} (S/N: ${d.sn})`;
            partDetailChartContainer.classList.remove('hidden');
            createPartDetailChart(currentDetailData, d.deliverDate);
            partDetailChartContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
        
        bars.filter(d => d.saleOrderDate && d.deliverDate).append("rect").attr("class", "main-chart-bar")
            .attr("x", d => xScale(d.saleOrderDate)).attr("y", 0).attr("width", d => Math.max(0, xScale(d.deliverDate) - xScale(d.saleOrderDate)))
            .attr("height", yScale.bandwidth()).attr("rx", 5).attr("ry", 5).attr("fill", "#4ade80")
            .on("mouseover", (event, d) => { tooltip.style("opacity", 0.9).html(`<strong>Customer:</strong> ${d.customer}<br/><strong>S/N:</strong> ${d.sn}<br/><strong>SO Nos:</strong> ${d.saleOrderNos}<br/><strong>Phase:</strong> Planned<br/><strong>Start:</strong> ${utcFormat(d.saleOrderDate)}<br/><strong>End:</strong> ${utcFormat(d.deliverDate)}<br/><em>Click to see part details</em>`).style("left", (event.pageX + 10) + "px").style("top", (event.pageY - 10) + "px"); })
            .on("mouseout", () => tooltip.style("opacity", 0)).on("click", handleBarClick);

        const delayedBars = bars.filter(d => d.isDelayed && d.deliverDate && d.etaDate && d.etaDate > d.deliverDate);
        delayedBars.append("rect").attr("class", "main-chart-bar")
            .attr("x", d => xScale(d.deliverDate)).attr("y", 0).attr("width", d => Math.max(0, xScale(d.etaDate) - xScale(d.deliverDate)))
            .attr("height", yScale.bandwidth()).attr("rx", 5).attr("ry", 5).attr("fill", "#f87171")
            .on("mouseover", (event, d) => { tooltip.style("opacity", 0.9).html(`<strong>Customer:</strong> ${d.customer}<br/><strong>S/N:</strong> ${d.sn}<br/><strong>SO Nos:</strong> ${d.saleOrderNos}<br/><strong>Phase:</strong> <span style='color:red;'>Delayed</span><br/><strong>Original End:</strong> ${utcFormat(d.deliverDate)}<br/><strong>New ETA:</strong> ${utcFormat(d.etaDate)}<br/><em>Click to see part details</em>`).style("left", (event.pageX + 10) + "px").style("top", (event.pageY - 10) + "px"); })
            .on("mouseout", () => tooltip.style("opacity", 0)).on("click", handleBarClick);
        
        delayedBars.append("text")
            .attr("class", "delay-text")
            .attr("x", d => xScale(d.deliverDate) + (xScale(d.etaDate) - xScale(d.deliverDate)) / 2)
            .attr("y", yScale.bandwidth() / 2)
            .text(d => {
                const weeks = d.maxDelay / (1000 * 60 * 60 * 24 * 7);
                return `${weeks.toFixed(1)}w`;
            });

        bars.append("text").attr("class", "bar-icon")
            .attr("x", d => xScale(d.saleOrderDate) + 15)
            .attr("y", yScale.bandwidth() / 2)
            .text("📄");

        bars.append("text").attr("class", "bar-icon")
            .attr("x", d => (d.isDelayed && d.etaDate && d.deliverDate && d.etaDate > d.deliverDate) ? xScale(d.etaDate) - 15 : xScale(d.deliverDate) - 15)
            .attr("y", yScale.bandwidth() / 2)
            .text("✈️");

    }
    
    function createPartDetailChart(detailData, deliverDate) {
        const chartId = "#part-detail-chart";
        d3.select(chartId).selectAll("*").remove();
        const margin = { top: 122, right: 40, bottom: 40, left: 400 };
        const containerWidth = document.querySelector(chartId).clientWidth;
        if (containerWidth <= 0) return;
        const chartHeight = Math.max(200, detailData.length * 50);
        const height = chartHeight + margin.top + margin.bottom;
        const width = containerWidth - margin.left - margin.right;
        
        const svgContainer = d3.select(chartId).append("svg").attr("width", containerWidth).attr("height", height);
        const svg = svgContainer.append("g").attr("transform", `translate(${margin.left},${margin.top})`);
        
        const yAxisGroup = svgContainer.append("g");
        const yHeader = svgContainer.append("g");
        
        const phases = [
            { start: 'OD', end: 'ETF', color: '#22c55e' },
            { start: 'ETF', end: 'ETD', color: '#8b5cf6' },
            { start: 'ETD', end: 'ETA', color: '#3b82f6' }
        ];
        
         // Dates are already pre-parsed, so we can filter for valid ones directly
        const allDates = detailData.flatMap(d => [d[excelHeaderMap.OD], d[excelHeaderMap.ETF], d[excelHeaderMap.ETD], d[excelHeaderMap.ETA]]).filter(Boolean);
        if (deliverDate) allDates.push(deliverDate);

        if (allDates.length === 0) {
            svg.append("text").attr("x", width/2).attr("y", chartHeight/2).attr("text-anchor", "middle").text("No date information available for these parts.");
            return;
        }
        const xScale = d3.scaleUtc().domain(d3.extent(allDates)).range([0, width]).nice();
        
        detailData.forEach((d, i) => d.partId = i);
        const yScale = d3.scaleBand().domain(detailData.map(d => d.partId)).range([0, chartHeight]).padding(0.4);

        drawSharedElements(svg, xScale, margin, chartHeight, width);
        
        const colDefs = [
            { name: "No.", width: 0.15, pos: 0 },
            { name: "Part Name", width: 0.55, pos: 0.15 },
            { name: "Progress", width: 0.30, pos: 0.70 }
        ];
        
        yHeader.attr("transform", `translate(0, 0)`);
        yHeader.append("rect").attr("x", 0).attr("y", 0).attr("width", margin.left).attr("height", margin.top - 2).attr("fill", "#fef9c3");
        
        colDefs.forEach(col => {
            yHeader.append("text").attr("class", "axis-header").attr("x", margin.left * (col.pos + col.width / 2)).attr("y", (margin.top - 2) / 2 + 5).attr("text-anchor", "middle").text(col.name);
            if (col.pos > 0) {
                yHeader.append("line").attr("x1", margin.left * col.pos).attr("x2", margin.left * col.pos).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af");
            }
        });
        yHeader.append("line").attr("x1", margin.left).attr("x2", margin.left).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af");

        const rowNumGroup = yAxisGroup.append("g").attr("transform", `translate(0, ${margin.top})`);
        const partNameGroup = yAxisGroup.append("g").attr("transform", `translate(${margin.left * colDefs[0].width}, ${margin.top})`);
        const progressGroup = yAxisGroup.append("g").attr("transform", `translate(${margin.left * (colDefs[0].width + colDefs[1].width)}, ${margin.top})`);

        rowNumGroup.selectAll("foreignObject").data(detailData).enter().append("foreignObject").attr("y", d => yScale(d.partId)).attr("width", margin.left * colDefs[0].width).attr("height", yScale.bandwidth()).append("xhtml:div").attr("class", "y-axis-label y-axis-row").html((d, i) => `<div class="row-number">${i + 1}</div>`);
        partNameGroup.selectAll("foreignObject").data(detailData).enter().append("foreignObject").attr("y", d => yScale(d.partId)).attr("width", margin.left * colDefs[1].width).attr("height", yScale.bandwidth()).append("xhtml:div").attr("class", "y-axis-label y-axis-row").html(d => `<div class="part-name-detail" title="${d[excelHeaderMap['Part Name']]}">${d[excelHeaderMap['Part Name']]}</div>`);
        progressGroup.selectAll("foreignObject").data(detailData).enter().append("foreignObject").attr("y", d => yScale(d.partId)).attr("width", margin.left * colDefs[2].width).attr("height", yScale.bandwidth()).append("xhtml:div").attr("class", "y-axis-label y-axis-row").html(d => `<div class="progress-detail" title="${d[excelHeaderMap['Progress']]}">${d[excelHeaderMap['Progress']]}</div>`);
        
        const bars = svg.selectAll(".bar-group").data(detailData).enter().append("g").attr("class", "bar-group").attr("transform", d => `translate(0, ${yScale(d.partId)})`);
        const barHeight = yScale.bandwidth();

        phases.forEach(phase => {
            const phaseGroup = bars.filter(d => d[excelHeaderMap[phase.start]] && d[excelHeaderMap[phase.end]]);
            phaseGroup.append("rect").attr("x", d => xScale(d[excelHeaderMap[phase.start]])).attr("y", 0).attr("width", d => Math.max(0, xScale(d[excelHeaderMap[phase.end]]) - xScale(d[excelHeaderMap[phase.start]]))).attr("height", barHeight).attr("fill", phase.color);
            let icon;
            if (phase.start === 'OD') icon = '💲';
            if (phase.start === 'ETF') icon = '📦';
            if (phase.start === 'ETD') {
                phaseGroup.append("text").attr("class", "milestone-icon").attr("x", d => xScale(d[excelHeaderMap[phase.start]])).attr("y", barHeight / 2).attr("dx", 10).style("pointer-events", "none").text(d => { const shipStatus = (d[excelHeaderMap['Ship']] || '').toUpperCase().trim(); if (shipStatus === 'A/F') return '✈️'; if (shipStatus === 'LAND' || shipStatus === 'ROAD') return '🚚'; return '🚢'; });
            } else { phaseGroup.append("text").attr("class", "milestone-icon").attr("x", d => xScale(d[excelHeaderMap[phase.start]])).attr("y", barHeight / 2).attr("dx", 10).style("pointer-events", "none").text(icon); }
            if (phase.end === 'ETA') { phaseGroup.append("text").attr("class", "milestone-icon").attr("x", d => xScale(d[excelHeaderMap[phase.end]])).attr("y", barHeight / 2).attr("dx", 10).style("pointer-events", "none").text('🏭'); }
        });

        if (deliverDate && deliverDate >= xScale.domain()[0] && deliverDate <= xScale.domain()[1]) {
            const deliverDateX = xScale(deliverDate);
            svg.append("line").attr("x1", deliverDateX).attr("x2", deliverDateX).attr("y1", -margin.top + 80).attr("y2", chartHeight).attr("stroke", "darkred").attr("stroke-width", 2.5);
            svg.append("text").attr("x", deliverDateX).attr("y", -margin.top + 95).attr("text-anchor", "middle").attr("fill", "darkred").style("font-size", "10px").style("font-weight", "bold").text("DELIVER DATE");
        }
    }
    
    function renderFullDetailGantt() {
        const customerName = customerFilterGanttDetail.value;

        if (!customerName) {
            ganttDetailTitleText.textContent = 'Select a Customer to Generate Report';
            ganttDetailTitleStats.innerHTML = '';
            d3.select("#gantt-detail-chart").selectAll("*").remove();
            printGanttDetailBtn.disabled = true;
            return;
        }

        // 1. Calculate stats for the whole customer
        ganttDetailTitleText.textContent = `Part Details Report for ${customerName}`;
        const allCustomerParts = fullRawData.filter(row => row[excelHeaderMap['Customer Name']] === customerName);

        if (allCustomerParts.length > 0) {
            let completedCount = 0;
            const completedProgress = ['completed', 'use stock', 'use local', 'n/a', 'cancelled', 'shipped'];
            allCustomerParts.forEach(part => {
                const solved = String(part[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
                const progress = String(part[excelHeaderMap['Progress']] || '').trim().toLowerCase();
                if (solved === 'YES' || completedProgress.includes(progress)) {
                    completedCount++;
                }
            });
            const totalParts = allCustomerParts.length;
            const completedPercent = totalParts > 0 ? ((completedCount / totalParts) * 100).toFixed(1) : "0.0";
            const incompletePercent = (100 - parseFloat(completedPercent)).toFixed(1);

            ganttDetailTitleStats.innerHTML = `
                <span class="text-green-800 bg-green-100 px-2 py-1 rounded">Completed: ${completedPercent}%</span>
                <span class="text-red-800 bg-red-100 px-2 py-1 rounded ml-2">Incomplete: ${incompletePercent}%</span>
            `;
        } else {
             ganttDetailTitleStats.innerHTML = '';
        }

        // 2. Filter data for chart rendering
        let chartData = [...allCustomerParts];
        const snTerm = searchSnDetail.value.toLowerCase().trim();
        if (snTerm) {
            chartData = chartData.filter(row => String(row[excelHeaderMap['S/N']] || '').toLowerCase().includes(snTerm));
        }
        const progressTerm = searchProgressDetail.value.toLowerCase().trim();
        if (progressTerm) {
            chartData = chartData.filter(row => String(row[excelHeaderMap['Progress']] || '').toLowerCase().includes(progressTerm));
        }

        // 3. Sort
        const getSortPriority = (row) => {
            const solved = String(row[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
            if (solved === 'NO' || solved === '??') return 1;
            return 2;
        };
        chartData.sort((a, b) => {
            const priorityA = getSortPriority(a);
            const priorityB = getSortPriority(b);
            if (priorityA !== priorityB) return priorityA - priorityB;

            const machineA = `${a[excelHeaderMap['M/C Type']] || ''}-${a[excelHeaderMap['S/N']] || ''}`;
            const machineB = `${b[excelHeaderMap['M/C Type']] || ''}-${b[excelHeaderMap['S/N']] || ''}`;
            if (machineA.localeCompare(machineB) !== 0) return machineA.localeCompare(machineB);
            
            const partNameA = a[excelHeaderMap['Part Name']] || '';
            const partNameB = b[excelHeaderMap['Part Name']] || '';
            return partNameA.localeCompare(partNameB);
        });

        // 4. Render
        createFullDetailGanttChart(chartData);
        printGanttDetailBtn.disabled = chartData.length === 0;
    }
    
    function printElementAsPng(targetId, fileName) {
        const node = document.getElementById(targetId);
        if (!node || !node.innerHTML.trim()) {
            alert("Nothing to print for this view.");
            return;
        }

        domtoimage.toPng(node, {
            width: node.scrollWidth,
            height: node.scrollHeight,
            bgcolor: '#ffffff'
        })
        .then(function (dataUrl) {
            const link = document.createElement('a');
            link.download = fileName;
            link.href = dataUrl;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        })
        .catch(function (error) {
            console.error(`Printing ${targetId} failed:`, error);
            alert('Could not print the report. See console for details.');
        });
    }

    function createFullDetailGanttChart(customerData) {
        const chartId = "#gantt-detail-chart";
        d3.select(chartId).selectAll("*").remove();
        if (!customerData || customerData.length === 0) return;

        const margin = { top: 122, right: 40, bottom: 40, left: 450 };
        const containerWidth = document.querySelector(chartId).clientWidth;
        if (containerWidth <= 0) return;

        const chartHeight = Math.max(300, customerData.length * 50);
        const height = chartHeight + margin.top + margin.bottom;
        const width = containerWidth - margin.left - margin.right;
        
        const svgContainer = d3.select(chartId).append("svg").attr("width", containerWidth).attr("height", height);
        const svg = svgContainer.append("g").attr("transform", `translate(${margin.left},${margin.top})`);
        const yAxisGroup = svgContainer.append("g");
        const yHeader = svgContainer.append("g");
        
        const phases = [ { start: 'OD', end: 'ETF', color: '#22c55e' }, { start: 'ETF', end: 'ETD', color: '#8b5cf6' }, { start: 'ETD', end: 'ETA', color: '#3b82f6' } ];
        
        const allDates = customerData.flatMap(d => [d[excelHeaderMap.OD], d[excelHeaderMap.ETF], d[excelHeaderMap.ETD], d[excelHeaderMap.ETA]]).filter(Boolean);

        if (allDates.length === 0) {
            svg.append("text").attr("x", width / 2).attr("y", chartHeight / 2).attr("text-anchor", "middle").text("No date information available for this customer.");
            return;
        }

        const xScale = d3.scaleUtc().domain(d3.extent(allDates)).range([0, width]).nice();
        customerData.forEach((d, i) => d.partId = i);
        const yScale = d3.scaleBand().domain(customerData.map(d => d.partId)).range([0, chartHeight]).padding(0.2);

        drawSharedElements(svg, xScale, margin, chartHeight, width);

        const colDefs = [
            { key: "no", name: "No.", width: 0.1, pos: 0 },
            { key: "partName", name: "Part Name", width: 0.45, pos: 0.1 },
            { key: "machineInfo", name: "Machine Info", width: 0.30, pos: 0.55 },
            { key: "progress", name: "Progress", width: 0.15, pos: 0.85 }
        ];
        
        yHeader.attr("transform", `translate(0, 0)`);
        yHeader.append("rect").attr("x", 0).attr("y", 0).attr("width", margin.left).attr("height", margin.top - 2).attr("fill", "#fef9c3");
        
        colDefs.forEach(col => {
            yHeader.append("text").attr("class", "axis-header").attr("x", margin.left * (col.pos + col.width / 2)).attr("y", (margin.top - 2) / 2 + 5).attr("text-anchor", "middle").text(col.name);
            if (col.pos > 0) yHeader.append("line").attr("x1", margin.left * col.pos).attr("x2", margin.left * col.pos).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af");
        });
         yHeader.append("line").attr("x1", margin.left).attr("x2", margin.left).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af");

        const machineGroups = [];
        if (customerData.length > 0) {
            let currentGroup = { mcType: customerData[0][excelHeaderMap['M/C Type']], sn: customerData[0][excelHeaderMap['S/N']], parts: [customerData[0]] };
            for (let i = 1; i < customerData.length; i++) {
                const part = customerData[i], mcType = part[excelHeaderMap['M/C Type']], sn = part[excelHeaderMap['S/N']];
                if (mcType === currentGroup.mcType && sn === currentGroup.sn) { currentGroup.parts.push(part); } 
                else { machineGroups.push(currentGroup); currentGroup = { mcType, sn, parts: [part] }; }
            }
            machineGroups.push(currentGroup);
        }

        const rowNumGroup = yAxisGroup.append("g").attr("transform", `translate(0, ${margin.top})`);
        const partNameGroup = yAxisGroup.append("g").attr("transform", `translate(${margin.left * colDefs[0].width}, ${margin.top})`);
        const machineInfoGroup = yAxisGroup.append("g").attr("transform", `translate(${margin.left * (colDefs[0].width + colDefs[1].width)}, ${margin.top})`);
        const progressGroup = yAxisGroup.append("g").attr("transform", `translate(${margin.left * (colDefs[0].width + colDefs[1].width + colDefs[2].width)}, ${margin.top})`);
        
        rowNumGroup.selectAll("foreignObject").data(customerData).enter().append("foreignObject").attr("y", d => yScale(d.partId)).attr("width", margin.left * colDefs[0].width).attr("height", yScale.bandwidth()).append("xhtml:div").attr("class", "y-axis-label y-axis-row").html((d, i) => `<div class="row-number">${i + 1}</div>`);
        partNameGroup.selectAll("foreignObject").data(customerData).enter().append("foreignObject").attr("y", d => yScale(d.partId)).attr("width", margin.left * colDefs[1].width).attr("height", yScale.bandwidth()).append("xhtml:div").attr("class", "y-axis-label y-axis-row").html(d => `<div class="part-name-full-detail" title="${d[excelHeaderMap['Part Name']]}">${d[excelHeaderMap['Part Name']]}</div>`);
        
        progressGroup.selectAll("foreignObject").data(customerData).enter().append("foreignObject").attr("y", d => yScale(d.partId)).attr("width", margin.left * colDefs[3].width).attr("height", yScale.bandwidth()).append("xhtml:div").attr("class", "y-axis-label y-axis-row").html(d => {
            const solved = String(d[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
            const progress = String(d[excelHeaderMap['Progress']] || '').trim().toLowerCase();
            const progressText = d[excelHeaderMap['Progress']] || 'N/A';
            let progressClass = '';
            if (solved === 'NO' || solved === '??') {
                progressClass = 'progress-highlight';
            } else if (progress === 'use stock' || progress === 'use local') {
                progressClass = 'progress-highlight-green';
            }
            return `<div class="progress-full-detail"><span class="${progressClass}" title="${progressText}">${progressText}</span></div>`;
        });

        machineInfoGroup.selectAll("foreignObject").data(machineGroups).enter().append("foreignObject")
            .attr("y", d => yScale(d.parts[0].partId))
            .attr("width", margin.left * colDefs[2].width)
            .attr("height", d => {
                const lastPart = d.parts[d.parts.length - 1];
                return (yScale(lastPart.partId) + yScale.bandwidth()) - yScale(d.parts[0].partId);
            })
            .append("xhtml:div").attr("class", "machine-info-merged-cell y-axis-row")
            .html(d => `<div><strong>M/C:</strong> ${d.mcType || 'N/A'}</div><div><strong>S/N:</strong> ${d.sn || 'N/A'}</div>`);
        
        const bars = svg.selectAll(".bar-group").data(customerData).enter().append("g").attr("class", "bar-group").attr("transform", d => `translate(0, ${yScale(d.partId)})`);
        const barHeight = yScale.bandwidth();

        phases.forEach(phase => {
            const phaseGroup = bars.filter(d => d[excelHeaderMap[phase.start]] && d[excelHeaderMap[phase.end]]);
            phaseGroup.append("rect").attr("x", d => xScale(d[excelHeaderMap[phase.start]])).attr("y", 0).attr("width", d => Math.max(0, xScale(d[excelHeaderMap[phase.end]]) - xScale(d[excelHeaderMap[phase.start]]))).attr("height", barHeight).attr("fill", phase.color);
            let icon;
            if (phase.start === 'OD') icon = '💲';
            if (phase.start === 'ETF') icon = '📦';
            if (phase.start === 'ETD') {
                phaseGroup.append("text").attr("class", "milestone-icon").attr("x", d => xScale(d[excelHeaderMap[phase.start]])).attr("y", barHeight / 2).attr("dx", 10).style("pointer-events", "none").text(d => { const shipStatus = (d[excelHeaderMap['Ship']] || '').toUpperCase().trim(); if (shipStatus === 'A/F') return '✈️'; if (shipStatus === 'LAND' || shipStatus === 'ROAD') return '🚚'; return '🚢'; });
            } else { phaseGroup.append("text").attr("class", "milestone-icon").attr("x", d => xScale(d[excelHeaderMap[phase.start]])).attr("y", barHeight / 2).attr("dx", 10).style("pointer-events", "none").text(icon); }
            if (phase.end === 'ETA') { phaseGroup.append("text").attr("class", "milestone-icon").attr("x", d => xScale(d[excelHeaderMap[phase.end]])).attr("y", barHeight / 2).attr("dx", 10).style("pointer-events", "none").text('🏭'); }
        });

        machineGroups.forEach(group => {
            const snData = fullCustomerData.find(d => d.sn === group.sn);
            if (snData && snData.deliverDate) {
                const deliverDate = snData.deliverDate;
                if (deliverDate && deliverDate >= xScale.domain()[0] && deliverDate <= xScale.domain()[1]) {
                    const deliverDateX = xScale(deliverDate);
                    svg.append("line").attr("x1", deliverDateX).attr("x2", deliverDateX).attr("y1", yScale(group.parts[0].partId)).attr("y2", yScale(group.parts[group.parts.length - 1].partId) + yScale.bandwidth()).attr("stroke", "darkred").attr("stroke-width", 2);
                    svg.append("text").attr("x", deliverDateX).attr("y", yScale(group.parts[0].partId) - 5).attr("text-anchor", "middle").attr("fill", "darkred").style("font-size", "10px").style("font-weight", "bold").text(`DD ${group.sn}`);
                }
            }
        });
    }

    function handleForecastBarClick(customerName) {
        miniCompletionTableContainer.classList.remove('hidden');
        renderMiniCompletionTable(customerName);
        miniCompletionTableContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function renderMonthlyForecastChart() {
        const year = forecastChartYearFilter.value;
        const month = forecastChartMonthFilter.value;
        miniCompletionTableContainer.classList.add('hidden');
        miniCompletionTableContainer.innerHTML = '';

        if (!year || !month) {
            monthlyForecastChartMessage.textContent = "Please select both a month and a year to generate the report.";
            monthlyForecastChartMessage.classList.remove('hidden');
            d3.select("#monthly-forecast-chart").selectAll("*").remove();
            monthlyForecastLegend.classList.add('hidden');
            printMonthlyForecastBtn.disabled = true;
            return;
        }

        let filteredData = fullRawData.filter(row => {
            // Logic is now simpler as dates are pre-parsed
            const forecastDate = row[excelHeaderMap['Forecast']];
            if (forecastDate) {
                return forecastDate.getUTCFullYear() == year && (forecastDate.getUTCMonth() + 1) == month;
            }
            return false;
        });

        if (filteredData.length === 0) {
             monthlyForecastChartMessage.textContent = "No data found for the selected month and year.";
             monthlyForecastChartMessage.classList.remove('hidden');
             d3.select("#monthly-forecast-chart").selectAll("*").remove();
             monthlyForecastLegend.classList.add('hidden');
             printMonthlyForecastBtn.disabled = true;
             return;
        }
        
        const groupedByCustomer = new Map();
        filteredData.forEach(row => {
            const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
            if (!groupedByCustomer.has(customer)) {
                groupedByCustomer.set(customer, {
                    customer,
                    parts: []
                });
            }
            groupedByCustomer.get(customer).parts.push(row);
        });

        const completedProgress = ['completed', 'use stock', 'use local', 'n/a', 'cancelled', 'shipped'];
        let chartData = Array.from(groupedByCustomer.values()).map(customerGroup => {
            const totalParts = customerGroup.parts.length;
            let completedParts = 0;
            customerGroup.parts.forEach(part => {
                 const solved = String(part[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
                 const progress = String(part[excelHeaderMap['Progress']] || '').trim().toLowerCase();
                 if (solved === 'YES' || completedProgress.includes(progress)) {
                     completedParts++;
                 }
            });
            const completedPercentage = totalParts > 0 ? (completedParts / totalParts) * 100 : 0;
            const incompletePercentage = 100 - completedPercentage;
            return {
                customer: customerGroup.customer,
                totalParts,
                completedParts,
                completedPercentage,
                incompletePercentage
            };
        });


        chartData.sort((a, b) => b.incompletePercentage - a.incompletePercentage);
        
        monthlyForecastChartMessage.classList.add('hidden');
        monthlyForecastLegend.classList.remove('hidden');
        const monthName = forecastChartMonthFilter.options[forecastChartMonthFilter.selectedIndex].text;
        monthlyForecastTitleText.textContent = `Monthly Forecast for ${monthName} ${year}`;
        printMonthlyForecastBtn.disabled = false;
        createMonthlyForecastGanttChart(chartData);
    }

    function createMonthlyForecastGanttChart(data) {
        const chartId = "#monthly-forecast-chart";
        d3.select(chartId).selectAll("*").remove();
        const margin = { top: 50, right: 40, bottom: 40, left: 250 };
        const containerWidth = document.querySelector(chartId).clientWidth;
        if (containerWidth <= 0) return;
        const chartHeight = Math.max(300, data.length * 60);
        const height = chartHeight + margin.top + margin.bottom;
        const width = containerWidth - margin.left - margin.right;
        const svgContainer = d3.select(chartId).append("svg").attr("width", containerWidth).attr("height", height);
        const svg = svgContainer.append("g").attr("transform", `translate(${margin.left},${margin.top})`);

        const xScale = d3.scaleLinear().domain([0, 100]).range([0, width]);
        const yScale = d3.scaleBand().domain(data.map(d => d.customer)).range([0, chartHeight]).padding(0.4);

        // Draw X Axis (Percentage)
        const xAxis = d3.axisTop(xScale).ticks(5).tickFormat(d => d + "%");
        svg.append("g").attr("class", "x-axis").call(xAxis);

        // Draw Y Axis Headers
        const yHeader = svgContainer.append("g").attr("transform", `translate(0, 0)`);
        yHeader.append("rect").attr("x", 0).attr("y", 0).attr("width", margin.left).attr("height", margin.top).attr("fill", "#fef9c3").attr("stroke", "#9ca3af");
        const yHeaderPositions = { no: 20, customer: 80 };
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.no).attr("y", margin.top / 2 + 5).text("No.");
        yHeader.append("line").attr("x1", yHeaderPositions.customer - 20).attr("x2", yHeaderPositions.customer - 20).attr("y1", 0).attr("y2", height).attr("stroke", "#9ca3af");
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.customer).attr("y", margin.top / 2 + 5).text("Customer Name");
        
        // Draw Y Axis Labels
        svgContainer.append("g").selectAll(".y-axis-label-group").data(data).enter().append("foreignObject")
            .attr("x", 0).attr("y", d => yScale(d.customer) + margin.top)
            .attr("width", margin.left).attr("height", yScale.bandwidth())
            .append("xhtml:div").attr("class", "y-axis-label y-axis-row")
            .html((d, i) => `<div class="row-number" style="width:25%">${i + 1}</div><div class="customer-name" style="width: 75%; border-right: none;" title="${d.customer}">${d.customer}</div>`);

        const bars = svg.selectAll(".bar-group").data(data).enter().append("g").attr("class", "bar-group").attr("transform", d => `translate(0, ${yScale(d.customer)})`);

        // Green Bar (Completed)
        bars.append("rect")
            .attr("x", 0)
            .attr("y", 0)
            .attr("width", d => xScale(d.completedPercentage))
            .attr("height", yScale.bandwidth())
            .attr("fill", "#4ade80")
            .style("cursor", "pointer")
             .on("mouseover", (event, d) => { tooltip.style("opacity", 0.9).html(`<strong>Customer:</strong> ${d.customer}<br/><strong>Completed:</strong> ${d.completedPercentage.toFixed(1)}%<br/>(${d.completedParts} of ${d.totalParts} parts)<br/><em>Click to see details</em>`).style("left", (event.pageX + 10) + "px").style("top", (event.pageY - 10) + "px"); })
            .on("mouseout", () => tooltip.style("opacity", 0))
            .on("click", (event, d) => handleForecastBarClick(d.customer));
            
        // Red Bar (Incomplete)
        bars.append("rect")
            .attr("x", d => xScale(d.completedPercentage))
            .attr("y", 0)
            .attr("width", d => xScale(100) - xScale(d.completedPercentage))
            .attr("height", yScale.bandwidth())
            .attr("fill", "#f87171")
            .style("cursor", "pointer")
            .on("mouseover", (event, d) => { tooltip.style("opacity", 0.9).html(`<strong>Customer:</strong> ${d.customer}<br/><strong>Incomplete:</strong> ${d.incompletePercentage.toFixed(1)}%<br/>(${d.totalParts - d.completedParts} of ${d.totalParts} parts)<br/><em>Click to see details</em>`).style("left", (event.pageX + 10) + "px").style("top", (event.pageY - 10) + "px"); })
            .on("mouseout", () => tooltip.style("opacity", 0))
            .on("click", (event, d) => handleForecastBarClick(d.customer));
        
        // Percentage Text
         bars.filter(d => d.completedPercentage > 10).append("text").attr("class", "percentage-text")
            .attr("x", d => xScale(d.completedPercentage) / 2)
            .attr("y", yScale.bandwidth() / 2)
            .text(d => `${d.completedPercentage.toFixed(0)}%`);

         bars.filter(d => d.incompletePercentage > 10).append("text").attr("class", "percentage-text")
            .attr("x", d => xScale(d.completedPercentage) + (xScale(100) - xScale(d.completedPercentage)) / 2)
            .attr("y", yScale.bandwidth() / 2)
            .text(d => `${d.incompletePercentage.toFixed(0)}%`);
    }


    function renderProgressTable() {
        if (!fullRawData.length) return;

        let data = getUniversallyFilteredRawData();

        const customerFilterValue = customerFilterProgress.value;
        const progressFilterValue = progressFilter.value;
        const sortKey = progressSort.value;

        if (customerFilterValue !== 'all') {
            data = data.filter(row => (row[excelHeaderMap['Customer Name']] || '') === customerFilterValue);
        }
        if (progressFilterValue !== 'all') {
            data = data.filter(row => (row[excelHeaderMap['Progress']] || '') === progressFilterValue);
        }

        if (sortKey === 'partName') {
            data.sort((a, b) => (a[excelHeaderMap['Part Name']] || '').localeCompare(b[excelHeaderMap['Part Name']] || ''));
        } else if (sortKey === 'progress') {
            data.sort((a, b) => (a[excelHeaderMap['Progress']] || '').localeCompare(b[excelHeaderMap['Progress']] || ''));
        }

        createProgressTable(data, sortKey);
    }

    function createProgressTable(data, sortKey) {
        progressTableBody.innerHTML = '';
        if (!excelHeaderMap['Progress']) return;

        let rowCounter = 1;

        if (sortKey === 'customer') {
            const groupedData = new Map();
            data.forEach(row => {
                const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
                const mcType = row[excelHeaderMap['M/C Type']] || 'N/A';
                const sn = row[excelHeaderMap['S/N']] || 'N/A';
                const mcKey = `${mcType} / ${sn}`;

                if (!groupedData.has(customer)) {
                    groupedData.set(customer, new Map());
                }
                if (!groupedData.get(customer).has(mcKey)) {
                    groupedData.get(customer).set(mcKey, []);
                }
                groupedData.get(customer).get(mcKey).push(row);
            });

            const sortedCustomers = [...groupedData.keys()].sort((a, b) => a.localeCompare(b));

            for (const customer of sortedCustomers) {
                const mcGroups = groupedData.get(customer);
                const sortedMcKeys = [...mcGroups.keys()].sort((a, b) => a.localeCompare(b));
                let firstMcRowForCustomer = true;
                const customerRowCount = Array.from(mcGroups.values()).reduce((sum, parts) => sum + parts.length, 0);

                for (const mcKey of sortedMcKeys) {
                    const parts = mcGroups.get(mcKey);
                    parts.sort((a, b) => (a[excelHeaderMap['Part Name']] || '').localeCompare(b[excelHeaderMap['Part Name']] || ''));
                    let firstPartRowForMc = true;
                    const mcRowCount = parts.length;

                    parts.forEach(part => {
                        const row = document.createElement('tr');
                        row.innerHTML = `<td>${rowCounter++}</td>`;
                        
                        if (firstMcRowForCustomer) {
                            const customerCell = document.createElement('td');
                            customerCell.textContent = customer;
                            customerCell.rowSpan = customerRowCount;
                            customerCell.classList.add('merged-cell');
                            row.appendChild(customerCell);
                            firstMcRowForCustomer = false;
                        }

                        if (firstPartRowForMc) {
                            const mcCell = document.createElement('td');
                            mcCell.textContent = mcKey;
                            mcCell.rowSpan = mcRowCount;
                            mcCell.classList.add('merged-cell');
                            row.appendChild(mcCell);
                            firstPartRowForMc = false;
                        }

                        const partNameCell = document.createElement('td');
                        partNameCell.textContent = part[excelHeaderMap['Part Name']] || 'N/A';
                        row.appendChild(partNameCell);

                        const currentProgress = String(part[excelHeaderMap['Progress']] || '').trim();
                        progressStatuses.forEach(status => {
                            const statusCell = document.createElement('td');
                            if (currentProgress.toLowerCase() === status.toLowerCase()) {
                                statusCell.innerHTML = '<div class="tick-mark">✔</div>';
                            }
                            row.appendChild(statusCell);
                        });
                        
                        progressTableBody.appendChild(row);
                    });
                }
            }
        } else {
            data.forEach(part => {
                const row = document.createElement('tr');
                
                const customer = part[excelHeaderMap['Customer Name']] || 'N/A';
                const mcType = part[excelHeaderMap['M/C Type']] || 'N/A';
                const sn = part[excelHeaderMap['S/N']] || 'N/A';
                const mcKey = `${mcType} / ${sn}`;
                const partName = part[excelHeaderMap['Part Name']] || 'N/A';

                row.innerHTML = `
                    <td>${rowCounter++}</td>
                    <td>${customer}</td>
                    <td>${mcKey}</td>
                    <td>${partName}</td>
                `;
                
                const currentProgress = String(part[excelHeaderMap['Progress']] || '').trim();
                progressStatuses.forEach(status => {
                    const statusCell = document.createElement('td');
                    if (currentProgress.toLowerCase() === status.toLowerCase()) {
                        statusCell.innerHTML = '<div class="tick-mark">✔</div>';
                    }
                    row.appendChild(statusCell);
                });
                
                progressTableBody.appendChild(row);
            });
        }
    }
    
    function renderArrivalAlarms() {
        if (!fullRawData.length) return;

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const oneDay = 24 * 60 * 60 * 1000;

        const excludedProgress = ['completed', 'use stock', 'use local', 'n/a'];
        const partsForAlarm = fullRawData.filter(row => {
            const progress = (row[excelHeaderMap['Progress']] || '').toLowerCase().trim();
            return !excludedProgress.includes(progress);
        });

        const arrived = partsForAlarm.filter(row => {
            const eta = row[excelHeaderMap['ETA']]; // Date is pre-parsed
            if (!eta) return false;
            const etaDateOnly = new Date(eta);
            etaDateOnly.setHours(0,0,0,0);
            return etaDateOnly <= today;
        });

        const arrivingIn1Week = partsForAlarm.filter(row => {
            const eta = row[excelHeaderMap['ETA']]; // Date is pre-parsed
            if (!eta) return false;
            const etaDateOnly = new Date(eta);
            etaDateOnly.setHours(0,0,0,0);
            const diffDays = Math.round((etaDateOnly - today) / oneDay);
            return diffDays > 0 && diffDays <= 7;
        });

        const arrivingIn2Weeks = partsForAlarm.filter(row => {
            const eta = row[excelHeaderMap['ETA']]; // Date is pre-parsed
            if (!eta) return false;
            const etaDateOnly = new Date(eta);
            etaDateOnly.setHours(0,0,0,0);
            const diffDays = Math.round((etaDateOnly - today) / oneDay);
            return diffDays > 7 && diffDays <= 14;
        });

        createAlarmGanttChart("#arrived-chart", arrived);
        createAlarmGanttChart("#arriving-1-week-chart", arrivingIn1Week);
        createAlarmGanttChart("#arriving-2-weeks-chart", arrivingIn2Weeks);
    }

    function createAlarmGanttChart(chartId, data) {
         d3.select(chartId).selectAll("*").remove();
         if (!data || data.length === 0) {
              d3.select(chartId).append("p").attr("class", "text-center text-gray-500 p-4").text("No parts in this category.");
              return;
         }
        
         const margin = { top: 122, right: 40, bottom: 40, left: 400 };
         const containerWidth = document.querySelector(chartId).clientWidth;
         if (containerWidth <= 0) return;
         const chartHeight = Math.max(200, data.length * 50);
         const height = chartHeight + margin.top + margin.bottom;
         const width = containerWidth - margin.left - margin.right;
         const svgContainer = d3.select(chartId).append("svg").attr("width", containerWidth).attr("height", height);
         const svg = svgContainer.append("g").attr("transform", `translate(${margin.left},${margin.top})`);
        
         const phases = [
              { start: 'OD', end: 'ETF', color: '#22c55e', icon: '💲' },
              { start: 'ETF', end: 'ETD', color: '#8b5cf6', icon: '📦' },
              { start: 'ETD', end: 'ETA', color: '#3b82f6', icon: '✈️' }
         ];
        
         const allDates = data.flatMap(d => [d[excelHeaderMap.OD], d[excelHeaderMap.ETF], d[excelHeaderMap.ETD], d[excelHeaderMap.ETA]]).filter(Boolean);
         if (allDates.length === 0) {
           svg.append("text").attr("x", width/2).attr("y", chartHeight/2).attr("text-anchor", "middle").text("No date information available.");
           return;
         }

         const xScale = d3.scaleUtc().domain(d3.extent(allDates)).range([0, width]).nice();
         data.forEach((d, i) => d.partId = i);
         const yScale = d3.scaleBand().domain(data.map(d => d.partId)).range([0, chartHeight]).padding(0.4);

         drawSharedElements(svg, xScale, margin, chartHeight, width);
        
        const yHeader = svgContainer.append("g").attr("transform", `translate(0, 0)`);
        yHeader.append("rect").attr("x", 0).attr("y", 0).attr("width", margin.left).attr("height", margin.top - 2).attr("fill", "#f3f4f6");
        
        const yHeaderPositions = { 
            customer: (margin.left * 0.0) + 10,
            partName: (margin.left * 0.4),
            prDraft:  (margin.left * 0.7)
        };
        
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.customer).attr("y", (margin.top - 2) / 2 + 5).text("Customer");
        
        yHeader.append("line").attr("x1", yHeaderPositions.partName - 5).attr("x2", yHeaderPositions.partName - 5).attr("y1", 0).attr("y2", height).attr("stroke", "#e5e7eb");
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.partName).attr("y", (margin.top - 2) / 2 + 5).text("Part Name");

        yHeader.append("line").attr("x1", yHeaderPositions.prDraft - 5).attr("x2", yHeaderPositions.prDraft - 5).attr("y1", 0).attr("y2", height).attr("stroke", "#e5e7eb");
        yHeader.append("text").attr("class", "axis-header").attr("x", yHeaderPositions.prDraft).attr("y", (margin.top - 2) / 2 + 5).text("PR Draft");

        svgContainer.append("g").selectAll(".y-axis-label-group").data(data).enter().append("foreignObject")
           .attr("x", 0).attr("y", d => yScale(d.partId) + margin.top)
           .attr("width", margin.left).attr("height", yScale.bandwidth())
           .append("xhtml:div").attr("class", "y-axis-label y-axis-row")
           .html(d => {
                const partName = d[excelHeaderMap['Part Name']] || 'N/A';
                const customer = d[excelHeaderMap['Customer Name']] || 'N/A';
                const prDraft = excelHeaderMap['PR Draft'] ? (d[excelHeaderMap['PR Draft']] || 'N/A') : 'N/A';
                return `<div class="customer-name" style="width: 40%; font-weight: normal; border-right: 1px solid #e5e7eb;" title="${customer}">${customer}</div>` +
                       `<div class="part-name-detail" style="width: 30%; border-right: 1px solid #e5e7eb;" title="${partName}">${partName}</div>` +
                       `<div class="pr-draft-detail" style="width: 30%; border-right: none; padding-left: 8px; align-items: center; display: flex;" title="${prDraft}">${prDraft}</div>`;
           });

        
         const bars = svg.selectAll(".bar-group").data(data).enter().append("g").attr("class", "bar-group").attr("transform", d => `translate(0, ${yScale(d.partId)})`);
         const barHeight = yScale.bandwidth();

         phases.forEach(phase => {
              const phaseGroup = bars.filter(d => d[excelHeaderMap[phase.start]] && d[excelHeaderMap[phase.end]]);
              
              phaseGroup.append("rect")
                   .attr("x", d => xScale(d[excelHeaderMap[phase.start]]))
                   .attr("y", 0)
                   .attr("width", d => {
                       const startDate = d[excelHeaderMap[phase.start]];
                       const endDate = d[excelHeaderMap[phase.end]];
                       return Math.max(0, xScale(endDate) - xScale(startDate));
                   })
                   .attr("height", barHeight)
                   .attr("fill", phase.color);
              
              phaseGroup.append("text")
                  .attr("class", "milestone-icon")
                  .attr("x", d => xScale(d[excelHeaderMap[phase.start]]))
                  .attr("y", barHeight / 2)
                  .attr("dx", 10)
                  .style("pointer-events", "none")
                  .text(phase.icon);

              if (phase.end === 'ETA') {
                   phaseGroup.append("text")
                        .attr("class", "milestone-icon")
                        .attr("x", d => xScale(d[excelHeaderMap[phase.end]]))
                        .attr("y", barHeight / 2)
                        .attr("dx", 10)
                        .style("pointer-events", "none")
                        .text('🏭');
              }
         });
    }

    function renderKpiReport() {
        if (!fullRawData.length) return;
        kpiTableBody.innerHTML = '';
        
        let data = getUniversallyFilteredRawData();

        const careOfFilter = kpiFilterCareOf.value;
        const progressFilter = kpiFilterProgress.value;

        const delayCauses = ['BOM Delay', 'Customer Delay', 'Drawing Delay', 'Engr. Delay', 'Long Lead Time', 'Delay'];
        
        let filteredData = data.filter(row => {
            const solved = String(row[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
            
            const careOfMatch = (careOfFilter === 'all') || ((row[excelHeaderMap['Care Of']] || '') === careOfFilter);
            const progressMatch = (progressFilter === 'all') || ((row[excelHeaderMap['Progress']] || '') === progressFilter);

            return (solved === 'NO' || solved === '??') && careOfMatch && progressMatch;
        });

        const groupedData = new Map();
        filteredData.forEach(row => {
            const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
            const mcType = row[excelHeaderMap['M/C Type']] || 'N/A';
            const sn = row[excelHeaderMap['S/N']] || 'N/A';
            const groupKey = `${customer}|${mcType}|${sn}`;

            if (!groupedData.has(groupKey)) {
                groupedData.set(groupKey, []);
            }
            groupedData.get(groupKey).push(row);
        });

        let itemCounter = 1;

        for (const [groupKey, parts] of groupedData.entries()) {
            const [customer, mcType, sn] = groupKey.split('|');
            let firstPart = true;

            parts.forEach(part => {
                const row = document.createElement('tr');
                
                const partName = part[excelHeaderMap['Part Name']] || 'N/A';
                let causeOfDelay = part[excelHeaderMap['Cause of Delay']] || '';
                const care = part[excelHeaderMap['Care Of']] || '';
                const progress = part[excelHeaderMap['Progress']] || '';
                const combinedInfo = `${care} | ${progress}`;

                row.innerHTML = `<td>${itemCounter++}</td>`;

                if (firstPart) {
                    row.innerHTML += `
                        <td rowspan="${parts.length}" class="merged-cell">${customer}</td>
                        <td rowspan="${parts.length}" class="merged-cell">${mcType}</td>
                        <td rowspan="${parts.length}" class="merged-cell">${sn}</td>
                    `;
                    firstPart = false;
                }

                row.innerHTML += `<td>${partName}</td>`;
                
                if (!causeOfDelay.trim() || !delayCauses.slice(0, -1).some(d => d.toLowerCase() === causeOfDelay.toLowerCase())) {
                    causeOfDelay = 'Delay';
                }

                delayCauses.forEach(cause => {
                    const cell = document.createElement('td');
                    cell.classList.add('kpi-delay-cell');
                    if (causeOfDelay.toLowerCase() === cause.toLowerCase()) {
                        cell.textContent = combinedInfo;
                    }
                    row.appendChild(cell);
                });
                
                kpiTableBody.appendChild(row);
            });
        }
    }

    function renderForecastTable() {
        if (!fullRawData.length) return;
        forecastTableBody.innerHTML = '';

        let data = getUniversallyFilteredRawData();
        const selectedMonth = forecastMonthFilter.value;

        let filteredData = data.filter(row => {
            const solved = String(row[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
            return solved === 'NO' || solved === '??';
        });

        if (selectedMonth !== 'all') {
            filteredData = filteredData.filter(row => {
                // Logic is now simpler as dates are pre-parsed
                const forecastDate = row[excelHeaderMap['Forecast']];
                if (forecastDate) {
                    const formattedMonth = forecastDate.toLocaleString('default', { month: 'short', year: 'numeric', timeZone: 'UTC' });
                    return formattedMonth === selectedMonth;
                }
                return false;
            });
        }
        
        const groupedData = new Map();
        filteredData.forEach(row => {
            // Logic is now simpler as dates are pre-parsed
            const forecastDate = row[excelHeaderMap['Forecast']];
            const formattedDate = forecastDate 
                ? forecastDate.toLocaleString('default', { month: 'short', year: 'numeric', timeZone: 'UTC' }) 
                : 'N/A';
            
            const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
            const mcType = row[excelHeaderMap['M/C Type']] || 'N/A';
            const sn = row[excelHeaderMap['S/N']] || 'N/A';
            
            const groupKey = `${formattedDate}|${customer}|${mcType}|${sn}`;

            if (!groupedData.has(groupKey)) {
                groupedData.set(groupKey, []);
            }
            groupedData.get(groupKey).push(row);
        });

        for (const [groupKey, parts] of groupedData.entries()) {
            const [forecast, customer, mcType, sn] = groupKey.split('|');
            let firstPart = true;

            parts.forEach(part => {
                const row = document.createElement('tr');
                
                if (firstPart) {
                    row.innerHTML = `
                        <td rowspan="${parts.length}" class="merged-cell">${forecast}</td>
                        <td rowspan="${parts.length}" class="merged-cell">${customer}</td>
                        <td rowspan="${parts.length}" class="merged-cell">${mcType}</td>
                        <td rowspan="${parts.length}" class="merged-cell">${sn}</td>
                    `;
                    firstPart = false;
                }

                const partName = part[excelHeaderMap['Part Name']] || 'N/A';
                const progress = part[excelHeaderMap['Progress']] || 'N/A';
                const remarks = excelHeaderMap['Remarks'] ? (part[excelHeaderMap['Remarks']] || '') : '';

                row.innerHTML += `
                    <td>${partName}</td>
                    <td>${progress}</td>
                    <td>${remarks}</td>
                `;
                
                forecastTableBody.appendChild(row);
            });
        }
    }

    function renderMiniCompletionTable(customerName) {
        const customerData = fullRawData.filter(row => row[excelHeaderMap['Customer Name']] === customerName);
         let tableHTML = `
            <h2 class="view-title">Completion Report for ${customerName}</h2>
            <table class="data-table">
                 <thead>
                     <tr>
                         <th>Item #</th>
                         <th>Forecast</th>
                         <th>M/C Type</th>
                         <th>S/N</th>
                         <th>Completed (%)</th>
                         <th>Incomplete (%)</th>
                         <th>Part Name</th>
                         <th>Progress</th>
                         <th>Remarks</th>
                     </tr>
                 </thead>
                 <tbody>`;

        if (customerData.length === 0) {
             tableHTML += `<tr><td colspan="9" class="text-center">No data available for this customer.</td></tr>`;
        } else {
             tableHTML += generateCompletionTableRows(customerData, true);
        }

         tableHTML += `</tbody></table>`;
         miniCompletionTableContainer.innerHTML = tableHTML;
    }

    function renderCompletionTable() {
        if (!fullRawData.length) return;

        let data = getUniversallyFilteredRawData();
        const selectedMonth = completionMonthFilter.value;
        const selectedPercentage = completionPercentageFilter.value;
        const selectedCustomers = Array.from(customerFilterCompletion.selectedOptions).map(opt => opt.value);
        const sortOrder = completionSort.value;

        if (selectedMonth !== 'all') {
            data = data.filter(row => {
                // Logic is now simpler as dates are pre-parsed
                const forecastDate = row[excelHeaderMap['Forecast']];
                if (forecastDate) {
                    const formattedMonth = forecastDate.toLocaleString('default', { month: 'short', year: 'numeric', timeZone: 'UTC' });
                    return formattedMonth === selectedMonth;
                }
                return false;
            });
        }

        if (selectedCustomers.length > 0) {
            data = data.filter(row => selectedCustomers.includes(row[excelHeaderMap['Customer Name']]));
        }

        const customerStats = new Map();
        data.forEach(row => {
            const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
            if (!customerStats.has(customer)) {
                customerStats.set(customer, { total: 0, completed: 0, completedPercent: 0 });
            }
            const stats = customerStats.get(customer);
            stats.total++;
            const solved = String(row[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
            const progress = String(row[excelHeaderMap['Progress']] || '').trim().toLowerCase();
            const completedProgress = ['completed', 'use stock', 'use local', 'n/a', 'cancelled', 'shipped'];
            if (solved === 'YES' || completedProgress.includes(progress)) {
                stats.completed++;
            }
        });
        
        customerStats.forEach(stats => {
            stats.completedPercent = stats.total > 0 ? (stats.completed / stats.total) * 100 : 0;
        });
        
        if (selectedPercentage !== 'all') {
            const customersToKeep = new Set();
            for (const [customer, stats] of customerStats.entries()) {
                 if (selectedPercentage === '<50' && stats.completedPercent < 50) customersToKeep.add(customer);
                 else if (selectedPercentage === '>=50' && stats.completedPercent >= 50 && stats.completedPercent < 100) customersToKeep.add(customer);
                 else if (selectedPercentage === '=100' && stats.completedPercent === 100) customersToKeep.add(customer);
            }
            data = data.filter(row => customersToKeep.has(row[excelHeaderMap['Customer Name']]));
        }
        
        completionTableBody.innerHTML = generateCompletionTableRows(data, false, customerStats, sortOrder);
    }
    
    function generateCompletionTableRows(data, isMini = false, externalCustomerStats = null, sortOrder = 'customer') {
        let tableRowsHTML = '';
        if (data.length === 0) return '<tr><td colspan="10" class="text-center">No data matches the selected filters.</td></tr>';

        const customerStats = externalCustomerStats || new Map();
         if (!externalCustomerStats) {
             data.forEach(row => {
                 const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
                 if (!customerStats.has(customer)) {
                     customerStats.set(customer, { total: 0, completed: 0, completedPercent: 0 });
                 }
                 const stats = customerStats.get(customer);
                 stats.total++;
                 const solved = String(row[excelHeaderMap['Solved ?']] || '').trim().toUpperCase();
                 const progress = String(row[excelHeaderMap['Progress']] || '').trim().toLowerCase();
                 const completedProgress = ['completed', 'use stock', 'use local', 'n/a', 'cancelled', 'shipped'];
                 if (solved === 'YES' || completedProgress.includes(progress)) {
                     stats.completed++;
                 }
             });
              customerStats.forEach(stats => {
                 stats.completedPercent = stats.total > 0 ? (stats.completed / stats.total) * 100 : 0;
             });
         }


        const getFormattedForecast = (part) => {
            // Logic is now simpler as dates are pre-parsed
            const forecastDate = part[excelHeaderMap['Forecast']];
            return forecastDate 
                ? forecastDate.toLocaleString('default', { month: 'short', year: 'numeric', timeZone: 'UTC' }) 
                : 'N/A';
        };
        
        const groupedByCustomer = new Map();
        data.forEach(row => {
            const customer = row[excelHeaderMap['Customer Name']] || 'N/A';
            if (!groupedByCustomer.has(customer)) {
                groupedByCustomer.set(customer, []);
            }
            groupedByCustomer.get(customer).push(row);
        });

        let itemCounter = 1;
        const highlightProgress = ['wait approve', 'wait po', 'wait pos', 'wait pr', 'wait qo', 'wip'];
        
        let sortedCustomers = [...groupedByCustomer.keys()];
        if (sortOrder === 'customer') {
             sortedCustomers.sort();
        } else if (sortOrder === 'completion_desc') {
            sortedCustomers.sort((a, b) => (customerStats.get(b)?.completedPercent || 0) - (customerStats.get(a)?.completedPercent || 0));
        } else if (sortOrder === 'completion_asc') {
             sortedCustomers.sort((a, b) => (customerStats.get(a)?.completedPercent || 0) - (customerStats.get(b)?.completedPercent || 0));
        }


        for (const customer of sortedCustomers) {
            const customerParts = groupedByCustomer.get(customer);

            const forecastSpans = {};
            if(customerParts.length > 0) {
                let currentForecast = getFormattedForecast(customerParts[0]);
                let spanCount = 0;
                let startIndex = 0;
                customerParts.forEach((part, index) => {
                    const forecast = getFormattedForecast(part);
                    if (forecast === currentForecast) {
                        spanCount++;
                    } else {
                        forecastSpans[startIndex] = spanCount;
                        startIndex = index;
                        spanCount = 1;
                        currentForecast = forecast;
                    }
                });
                forecastSpans[startIndex] = spanCount;
            }
            
            const machineSpans = {};
             if(customerParts.length > 0) {
                let currentMachineKey = `${customerParts[0][excelHeaderMap['M/C Type']] || 'N/A'}|${customerParts[0][excelHeaderMap['S/N']] || 'N/A'}`;
                let spanCount = 0;
                let startIndex = 0;
                customerParts.forEach((part, index) => {
                    const mcType = part[excelHeaderMap['M/C Type']] || 'N/A';
                    const sn = part[excelHeaderMap['S/N']] || 'N/A';
                    const machineKey = `${mcType}|${sn}`;
                    if (machineKey === currentMachineKey) {
                        spanCount++;
                    } else {
                        machineSpans[startIndex] = { mcType: customerParts[startIndex][excelHeaderMap['M/C Type']] || 'N/A', sn: customerParts[startIndex][excelHeaderMap['S/N']] || 'N/A', span: spanCount };
                        startIndex = index;
                        spanCount = 1;
                        currentMachineKey = machineKey;
                    }
                });
                machineSpans[startIndex] = { mcType: customerParts[startIndex][excelHeaderMap['M/C Type']] || 'N/A', sn: customerParts[startIndex][excelHeaderMap['S/N']] || 'N/A', span: spanCount };
            }


            customerParts.forEach((part, index) => {
                tableRowsHTML += `<tr>`;
                tableRowsHTML += `<td>${itemCounter++}</td>`;

                if (forecastSpans[index]) {
                    tableRowsHTML += `<td rowspan="${forecastSpans[index]}" class="merged-cell">${getFormattedForecast(part)}</td>`;
                }

                if (index === 0) {
                    const customerRowCount = customerParts.length;
                    const stats = customerStats.get(customer);
                    const completedPercent = stats.completedPercent.toFixed(1);
                    const incompletePercent = (100 - stats.completedPercent).toFixed(1);
                    
                    if (!isMini) {
                        tableRowsHTML += `<td rowspan="${customerRowCount}" class="merged-cell">${customer}</td>`;
                    }
                    
                    if (machineSpans[index]) {
                        const { mcType, sn, span } = machineSpans[index];
                        tableRowsHTML += `<td rowspan="${span}" class="merged-cell">${mcType}</td>`;
                        tableRowsHTML += `<td rowspan="${span}" class="merged-cell">${sn}</td>`;
                    }
                    
                    tableRowsHTML += `<td rowspan="${customerRowCount}" class="merged-cell">${completedPercent}%</td>`;
                    tableRowsHTML += `<td rowspan="${customerRowCount}" class="merged-cell">${incompletePercent}%</td>`;
                } else {
                     if (machineSpans[index]) {
                        const { mcType, sn, span } = machineSpans[index];
                        tableRowsHTML += `<td rowspan="${span}" class="merged-cell">${mcType}</td>`;
                        tableRowsHTML += `<td rowspan="${span}" class="merged-cell">${sn}</td>`;
                    }
                }
                
                const partName = part[excelHeaderMap['Part Name']] || 'N/A';
                const progress = part[excelHeaderMap['Progress']] || 'N/A';
                let remarks = excelHeaderMap['Remarks'] ? (part[excelHeaderMap['Remarks']] || '') : '';
                const progressClass = highlightProgress.includes(progress.toLowerCase()) ? 'class="progress-highlight"' : '';

                if (progress.toLowerCase() === 'wip') {
                    const etaDate = part[excelHeaderMap['ETA']];
                    const prDraft = excelHeaderMap['PR Draft'] ? (part[excelHeaderMap['PR Draft']] || 'N/A') : 'N/A';
                    let etaString = 'N/A';
                    if (etaDate instanceof Date && !isNaN(etaDate)) {
                        const day = ('0' + etaDate.getUTCDate()).slice(-2);
                        const month = ('0' + (etaDate.getUTCMonth() + 1)).slice(-2);
                        const year = etaDate.getUTCFullYear().toString().slice(-2);
                        etaString = `${day}/${month}/${year}`;
                    }
                    remarks += ` <span class="progress-highlight">(ETA:${etaString} PR Draft: ${prDraft})</span>`;
                }

                tableRowsHTML += `
                    <td>${partName}</td>
                    <td ${progressClass}>${progress}</td>
                    <td>${remarks}</td>
                `;
                tableRowsHTML += `</tr>`;
            });
        }
        return tableRowsHTML;
    }


    function showError(message) {
        messageText.innerHTML = message;
        messageArea.classList.remove('hidden');
    }

    const dateEl = document.getElementById('current-date');
    const timeEl = document.getElementById('current-time');
    function updateTime() {
        const now = new Date();
        dateEl.textContent = now.toLocaleDateString(undefined, { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' });
        timeEl.textContent = now.toLocaleTimeString();
    }
    setInterval(updateTime, 1000);
    updateTime();

});