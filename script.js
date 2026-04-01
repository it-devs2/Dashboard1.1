/**
 * ==== การเชื่อมต่อ GOOGLE SHEETS ====
 * ให้คุณใส่ URL ของ Web App จาก Google Apps Script ตรงนี้
 */
const GOOGLE_APP_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzo_H47zLlED0QoCcM6OLM5PLG6EGRozcWiLjJWn8iwGJ5ayZhJrW6AQQtReed7R_Pv/exec';


// ตัวแปรเก็บข้อมูลทั้งหมดจาก Google Sheets
let allData = [];
let currentFilteredData = [];
// ตัวแปรเก็บกราฟ
let donutChart;
let barChart;

// Format numbers as Thai Baht currency
const formatCurrency = (number) => {
    return new Intl.NumberFormat('th-TH', {
        style: 'currency',
        currency: 'THB',
        minimumFractionDigits: 2
    }).format(number);
};

// Number counter animation function
const animateValue = (obj, start, end, duration, isCurrency = false) => {
    let startTimestamp = null;
    const step = (timestamp) => {
        if (!startTimestamp) startTimestamp = timestamp;
        const progress = Math.min((timestamp - startTimestamp) / duration, 1);
        const ease = 1 - Math.pow(1 - progress, 4); // easeOutQuart
        const currentVal = ease * (end - start) + start;

        if (isCurrency) {
            obj.innerText = formatCurrency(currentVal);
        } else {
            obj.innerText = `คิดเป็น ${currentVal.toFixed(2)}%`;
        }

        if (progress < 1) {
            window.requestAnimationFrame(step);
        } else {
            if (isCurrency) obj.innerText = formatCurrency(end);
            else obj.innerText = `คิดเป็น ${end.toFixed(2)}%`;
        }
    };
    window.requestAnimationFrame(step);
};

// DOM Elements
const paymentStatusFilter = document.getElementById('paymentStatusFilter');
const categoryFilter = document.getElementById('categoryFilter');
const monthFilter = document.getElementById('monthFilter');
const dayFilter = document.getElementById('dayFilter');
const yearFilter = document.getElementById('yearFilter');
const refreshBtn = document.getElementById('refreshBtn');
const loading = document.getElementById('loading');

const exportPdfBtn = document.getElementById('exportPdfBtn');

const totalAmountEl = document.getElementById('totalAmount');
const overdueAmountEl = document.getElementById('overdueAmount');
const ontimeAmountEl = document.getElementById('ontimeAmount');
const notdueAmountEl = document.getElementById('notdueAmount');

const totalPercentEl = document.getElementById('totalPercent');
const overduePercentEl = document.getElementById('overduePercent');
const ontimePercentEl = document.getElementById('ontimePercent');
const notduePercentEl = document.getElementById('notduePercent');

const nodateAmountEl = document.getElementById('nodateAmount');
const earlyAmountEl = document.getElementById('earlyAmount');
const nodatePercentEl = document.getElementById('nodatePercent');
const earlyPercentEl = document.getElementById('earlyPercent');

const pendingAmountEl = document.getElementById('pendingAmount');
const pendingPercentEl = document.getElementById('pendingPercent');

// Initialize the dashboard
const init = async () => {
    setupEventListeners();

    const now = new Date();
    yearFilter.value = now.getFullYear().toString();

    initCharts();

    // Populate dayFilter 1-31
    if (dayFilter) {
        for (let i = 1; i <= 31; i++) {
            const opt = document.createElement('option');
            opt.value = i.toString().padStart(2, '0');
            opt.textContent = i.toString();
            dayFilter.appendChild(opt);
        }
    }

    if (GOOGLE_APP_SCRIPT_URL === 'YOUR_WEB_APP_URL_HERE') {
        document.getElementById('setupModal').classList.remove('hidden');
        loadMockData();
    } else {
        await fetchData();
    }
};

// Setup Listeners
const setupEventListeners = () => {
    paymentStatusFilter.addEventListener('change', updateDashboard);
    categoryFilter.addEventListener('change', updateDashboard);
    dayFilter.addEventListener('change', updateDashboard);
    monthFilter.addEventListener('change', updateDashboard);
    yearFilter.addEventListener('change', updateDashboard);

    refreshBtn.addEventListener('click', async () => {
        if (GOOGLE_APP_SCRIPT_URL === 'YOUR_WEB_APP_URL_HERE') {
            alert('กรุณาใส่ Web App URL ของคุณในไฟล์ script.js ก่อนครับ');
        } else {
            await fetchData();
        }
    });

    // Modal Close
    const modal = document.getElementById('setupModal');
    const closeBtn = document.querySelector('.close-btn');
    const closeBtn2 = document.getElementById('closeModalBtn');

    closeBtn.addEventListener('click', () => modal.classList.add('hidden'));
    closeBtn2.addEventListener('click', () => modal.classList.add('hidden'));

    // Details Modal Setup
    const detailsModal = document.getElementById('detailsModal');
    const closeDetailsBtn = document.querySelector('.close-details-btn');
    const closeDetailsModalBtn = document.getElementById('closeDetailsModalBtn');

    closeDetailsBtn.addEventListener('click', () => detailsModal.classList.add('hidden'));
    closeDetailsModalBtn.addEventListener('click', () => detailsModal.classList.add('hidden'));

    // PDF Export
    if (exportPdfBtn) {
        exportPdfBtn.addEventListener('click', exportToPDF);
    }

    // Details Buttons
    document.querySelectorAll('.details-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const statusType = e.currentTarget.getAttribute('data-status');
            openDetailsModal(statusType);
        });
    });

    window.addEventListener('click', (e) => {
        if (e.target === modal) modal.classList.add('hidden');
        if (e.target === detailsModal) detailsModal.classList.add('hidden');
    });

};

// Open details Modal and populate table
const openDetailsModal = (type) => {
    const detailsModal = document.getElementById('detailsModal');
    const detailsModalTitle = document.getElementById('detailsModalTitle');
    const detailsTableBody = document.getElementById('detailsTableBody');

    // Define status mappings
    const statusMap = {
        'overdue': { text: 'จ่ายเกินกำหนด', key: 'เกินกำหนด' },
        'ontime': { text: 'จ่ายตรงดิว', key: 'ตรงดิว' },
        'notdue': { text: 'ยังไม่ถึงกำหนด', key: 'ยังไม่ถึงกำหนด' },
        'nodate': { text: 'ยังไม่กำหนดวันจ่าย', key: 'ยังไม่กำหนดวันจ่าย' },
        'early': { text: 'จ่ายก่อนกำหนด', key: 'จ่ายก่อนกำหนด' },
        'pending': { text: 'เกินกำหนด (รอพิจารณา)', key: 'เกินกำหนด (รอพิจารณา)' }
    };

    const config = statusMap[type];
    if (!config) return;

    const titleText = `ประเภทรายงาน: ${config.text}`;
    detailsModalTitle.innerText = titleText;

    // ตั้งค่าสำหรับหัวข้อที่จะพิมพ์ใน PDF (แยกจากหน้าจอ)
    const printHeader = document.getElementById('printReportHeader');
    if (printHeader) printHeader.innerText = titleText;

    // Filter and sort data based on current context
    const items = currentFilteredData.filter(item => {
        const s = (item.status || '').toString().trim();
        return s.includes(config.key);
    }).sort((a, b) => (Number(b.amount) || 0) - (Number(a.amount) || 0));

    detailsTableBody.innerHTML = '';

    if (items.length === 0) {
        detailsTableBody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding: 32px; color: var(--text-muted);">ไม่มีข้อมูลสำหรับตัวกรองนี้</td></tr>`;
    } else {
        let totalSum = 0;
        items.forEach(item => {
            const amount = Number(item.amount) || 0;
            totalSum += amount;

            const tr = document.createElement('tr');
            const dueDateStr = [item.dayDue, item.monthDue, item.yearDue].filter(Boolean).join(' ') || '-';
            tr.innerHTML = `
                <td style="font-weight: 500; color: var(--accent-primary); font-family: monospace;">${item.docNo || '-'}</td>
                <td style="font-weight: 500;">${item.creditor || '-'}</td>
                <td style="color: var(--text-muted);">${item.description || '-'}</td>
                <td><span style="background: rgba(255,255,255,0.1); padding: 4px 8px; border-radius: 4px; font-size: 12px; white-space: nowrap;">${item.category || '-'}</span></td>
                <td style="white-space: nowrap; font-size: 12px; color: var(--text-muted);">${dueDateStr}</td>
                <td style="text-align: right; color: var(--color-total); font-weight: 600; white-space: nowrap;">${formatCurrency(amount)}</td>
            `;
            detailsTableBody.appendChild(tr);
        });

        // Add Total Summary Row to tfoot
        const detailsTableFooter = document.getElementById('detailsTableFooter');
        detailsTableFooter.innerHTML = ''; // Clear previous

        const totalTr = document.createElement('tr');
        totalTr.className = 'total-row-summary';
        totalTr.innerHTML = `
            <td colspan="5" class="total-label">ยอดรวมทั้งหมด (Total):</td>
            <td class="total-amount-val">${formatCurrency(totalSum)}</td>
        `;
        detailsTableFooter.appendChild(totalTr);
    }

    detailsModal.classList.remove('hidden');
};

// Initializing empty charts
const initCharts = () => {
    // 1. Donut Chart (Status)
    const ctxStatus = document.getElementById('statusChart').getContext('2d');

    // Shared styling properties
    Chart.defaults.color = '#8e8e9e';
    Chart.defaults.font.family = "'Prompt', sans-serif";

    donutChart = new Chart(ctxStatus, {
        type: 'doughnut',
        data: {
            labels: ['จ่ายเกินกำหนด', 'จ่ายตรงดิว', 'ยังไม่ถึงกำหนด', 'ยังไม่กำหนดวันจ่าย', 'จ่ายก่อนกำหนด', 'เกินกำหนด (รอพิจารณา)'],
            datasets: [{
                data: [],
                backgroundColor: ['#ef4444', '#10b981', '#3b82f6', '#f59e0b', '#a855f7', '#f97316'],
                borderWidth: 0,
                hoverOffset: 12
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '70%',
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: 'rgba(255,255,255,0.7)',
                        usePointStyle: true,
                        padding: 20,
                        font: { family: 'Inter, Prompt, sans-serif', size: 12 }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 15, 25, 0.95)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    titleFont: { size: 14, weight: 'bold', family: 'Inter, Prompt, sans-serif' },
                    bodyFont: { size: 13, family: 'Inter, Prompt, sans-serif' },
                    padding: 12,
                    cornerRadius: 10,
                    borderColor: 'rgba(99, 102, 241, 0.3)',
                    borderWidth: 1,
                    displayColors: true,
                    boxPadding: 6,
                    callbacks: {
                        label: function (context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const val = context.raw;
                            const pct = total === 0 ? 0 : ((val / total) * 100).toFixed(2);
                            return [
                                ` ประเภท: ${context.label}`,
                                ` ยอดรวม: ${formatCurrency(val)}`,
                                ` สัดส่วน: ${pct}%`
                            ];
                        }
                    }
                }
            }
        }
    });

    // 2. Bar Chart (Top Expenses by Creditor/Category)
    const ctxCategory = document.getElementById('categoryChart').getContext('2d');

    // Custom inline plugin for bar value labels
    const barValueLabels = {
        id: 'barValueLabels',
        afterDatasetsDraw(chart) {
            const { ctx, data } = chart;
            data.datasets.forEach((dataset, datasetIndex) => {
                const meta = chart.getDatasetMeta(datasetIndex);
                meta.data.forEach((bar, index) => {
                    const value = dataset.data[index];
                    if (value === undefined || value === null || value === 0) return;

                    let label;
                    if (value >= 1000000) label = '\u0e3f' + (value / 1000000).toFixed(2) + 'M';
                    else if (value >= 1000) label = '\u0e3f' + (value / 1000).toFixed(1) + 'K';
                    else label = '\u0e3f' + value.toLocaleString();

                    const x = bar.x;
                    const y = bar.y - 12;

                    ctx.save();
                    ctx.font = 'bold 11px Inter, sans-serif';
                    const textWidth = ctx.measureText(label).width;
                    const padX = 8, padY = 4;
                    const pillW = textWidth + padX * 2;
                    const pillH = 22;
                    const pillX = x - pillW / 2;
                    const pillY = y - pillH;

                    ctx.beginPath();
                    ctx.roundRect(pillX, pillY, pillW, pillH, 6);
                    ctx.fillStyle = 'rgba(139, 92, 246, 0.9)';
                    ctx.shadowColor = 'rgba(139, 92, 246, 0.5)';
                    ctx.shadowBlur = 8;
                    ctx.fill();
                    ctx.shadowBlur = 0;

                    ctx.fillStyle = '#ffffff';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillText(label, x, pillY + pillH / 2);
                    ctx.restore();
                });
            });
        }
    };

    barChart = new Chart(ctxCategory, {
        type: 'bar',
        data: {
            labels: [],
            datasets: [{
                label: '\u0e22\u0e2d\u0e14\u0e43\u0e0a\u0e49\u0e08\u0e48\u0e32\u0e22 (\u0e1a\u0e32\u0e17)',
                data: [],
                backgroundColor: function (context) {
                    const chart = context.chart;
                    const { ctx, chartArea } = chart;
                    if (!chartArea) return 'rgba(99, 102, 241, 0.8)';
                    const gradient = ctx.createLinearGradient(0, chartArea.top, 0, chartArea.bottom);
                    gradient.addColorStop(0, 'rgba(167, 139, 250, 0.95)');
                    gradient.addColorStop(1, 'rgba(99, 102, 241, 0.55)');
                    return gradient;
                },
                borderRadius: 8,
                borderSkipped: false,
                hoverBackgroundColor: 'rgba(192, 132, 252, 1)'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 44 } },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: {
                        color: 'rgba(255,255,255,0.45)',
                        callback: function (value) {
                            if (value >= 1000000) return '\u0e3f' + (value / 1000000).toFixed(1) + 'M';
                            if (value >= 1000) return '\u0e3f' + (value / 1000).toFixed(0) + 'K';
                            return '\u0e3f' + value;
                        }
                    }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: 'rgba(255,255,255,0.6)', maxRotation: 30 }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15, 15, 25, 0.95)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    titleFont: { size: 14, weight: 'bold', family: 'Inter, Prompt, sans-serif' },
                    bodyFont: { size: 13, family: 'Inter, Prompt, sans-serif' },
                    padding: 12,
                    cornerRadius: 10,
                    borderColor: 'rgba(99, 102, 241, 0.3)',
                    borderWidth: 1,
                    callbacks: {
                        title: function (context) {
                            return '👤 ชื่อเจ้าหนี้: ' + context[0].label;
                        },
                        label: function (context) {
                            const dataset = context.dataset;
                            const statusStr = dataset.statusData ? dataset.statusData[context.dataIndex] : 'ไม่ทราบกลุ่ม';
                            return [
                                ' 📦 ข้อมูลจากกลุ่ม: ' + statusStr,
                                ' 💰 ยอดเงิน: ' + formatCurrency(context.raw)
                            ];
                        }
                    }
                }
            }
        },
        plugins: [barValueLabels]
    });
};

// Fetch data from Google Sheets API
const fetchData = async () => {
    loading.classList.remove('hidden');
    try {
        const response = await fetch(GOOGLE_APP_SCRIPT_URL);
        const result = await response.json();

        if (result.status === 'success') {
            allData = result.data;
            updateDashboard();
        } else {
            console.error('API Error:', result.message);
            alert('เกิดข้อผิดพลาดในการดึงข้อมูลจาก Google Sheets: ' + result.message);
        }
    } catch (error) {
        console.error('Fetch Error:', error);
        alert('เกิดข้อผิดพลาดในการเชื่อมต่อ กรุณาตรวจสอบ URL ของ Web App');
    } finally {
        loading.classList.add('hidden');
    }
};


// Update Dashboard View based on selected filters
const updateDashboard = () => {
    const selectedPaymentStatus = paymentStatusFilter.value;
    const selectedCategory = categoryFilter.value;
    const selectedDay = dayFilter.value;
    const selectedMonth = monthFilter.value;
    const selectedYear = yearFilter.value;

    // Filter data
    currentFilteredData = allData.filter(item => {
        let matchPaymentStatus = selectedPaymentStatus === 'all' || (item.paymentStatus && item.paymentStatus.toString().includes(selectedPaymentStatus));
        let matchCategory = selectedCategory === 'all' || (item.category && item.category.toString().includes(selectedCategory));
        let matchDay = selectedDay === 'all' || (item.dayDue && parseInt(item.dayDue) === parseInt(selectedDay));
        let matchMonth = selectedMonth === 'all' || (item.monthDue && item.monthDue.toString() === selectedMonth);
        let matchYear = selectedYear === 'all' || (item.yearDue && parseInt(item.yearDue) === parseInt(selectedYear));
        return matchPaymentStatus && matchCategory && matchDay && matchMonth && matchYear;
    });

    // Calculate Summary numbers
    let total = 0, overdue = 0, ontime = 0, notdue = 0, nodate = 0, early = 0, pending = 0;

    // Temporary object to group data for bar chart (By Creditor - ชื่อเจ้าหนี้การค้า)
    const creditorSummary = {};

    currentFilteredData.forEach(item => {
        // Convert string to number just in case
        const amount = Number(item.amount) || 0;

        // Sum total directly to ensure 100% accuracy with Google Sheets
        total += amount;

        // Status calculation (N column = ความเร่งด่วน: เกินกำหนด/ตรงดิว/ยังไม่ถึงกำหนด)
        const statusStr = (item.status || '').toString().trim();

        // Count strictly correctly to avoid overlap bugs
        if (statusStr.includes('เกินกำหนด (รอพิจารณา)')) {
            pending += amount;
        } else if (statusStr.includes('เกินกำหนด')) {
            overdue += amount;
        } else if (statusStr.includes('ตรงดิว')) {
            ontime += amount;
        } else if (statusStr.includes('ยังไม่ถึงกำหนด')) {
            notdue += amount;
        } else if (statusStr.includes('ยังไม่กำหนดวันจ่าย')) {
            nodate += amount;
        } else if (statusStr.includes('จ่ายก่อนกำหนด')) {
            early += amount;
        }

        // Group by creditor and track their statuses for the bar chart
        const creditor = item.creditor ? item.creditor : 'ไม่ระบุชื่อ';
        if (!creditorSummary[creditor]) {
            creditorSummary[creditor] = { amount: 0, statuses: new Set() };
        }
        creditorSummary[creditor].amount += amount;

        // Map status strings to short readable box names
        let boxName = "อื่น ๆ";
        if (statusStr.includes('เกินกำหนด (รอพิจารณา)')) boxName = 'เกินกำหนด (รอพิจารณา)';
        else if (statusStr.includes('เกินกำหนด')) boxName = 'จ่ายเกินกำหนด';
        else if (statusStr.includes('ตรงดิว')) boxName = 'จ่ายตรงดิว';
        else if (statusStr.includes('ยังไม่ถึงกำหนด')) boxName = 'ยังไม่ถึงกำหนด';
        else if (statusStr.includes('ยังไม่กำหนดวันจ่าย')) boxName = 'ยังไม่กำหนดวันจ่าย';
        else if (statusStr.includes('จ่ายก่อนกำหนด')) boxName = 'จ่ายก่อนกำหนด';

        creditorSummary[creditor].statuses.add(boxName);
    });

    // Total is already calculated in the loop above to include all items correctly
    // total = overdue + ontime + notdue + nodate + early + pending;

    // Update Text Elements with Counting Animation
    animateValue(totalAmountEl, 0, total, 1200, true);
    animateValue(overdueAmountEl, 0, overdue, 1200, true);
    animateValue(ontimeAmountEl, 0, ontime, 1200, true);
    animateValue(notdueAmountEl, 0, notdue, 1200, true);
    animateValue(nodateAmountEl, 0, nodate, 1200, true);
    animateValue(earlyAmountEl, 0, early, 1200, true);
    animateValue(pendingAmountEl, 0, pending, 1200, true);

    // Update Percentages
    totalPercentEl.innerText = `คิดเป็น 100.00%`;

    const overduePct = total === 0 ? 0 : (overdue / total) * 100;
    const ontimePct = total === 0 ? 0 : (ontime / total) * 100;
    const notduePct = total === 0 ? 0 : (notdue / total) * 100;
    const nodatePct = total === 0 ? 0 : (nodate / total) * 100;
    const earlyPct = total === 0 ? 0 : (early / total) * 100;
    const pendingPct = total === 0 ? 0 : (pending / total) * 100;

    animateValue(overduePercentEl, 0, overduePct, 1200, false);
    animateValue(ontimePercentEl, 0, ontimePct, 1200, false);
    animateValue(notduePercentEl, 0, notduePct, 1200, false);
    animateValue(nodatePercentEl, 0, nodatePct, 1200, false);
    animateValue(earlyPercentEl, 0, earlyPct, 1200, false);
    animateValue(pendingPercentEl, 0, pendingPct, 1200, false);

    // Update Donut Chart
    donutChart.data.datasets[0].data = [overdue, ontime, notdue, nodate, early, pending];
    donutChart.update();

    // Prepare Bar Chart Data (Sort by Highest Amount & take top 10)
    const sortedCreditors = Object.entries(creditorSummary)
        .sort((a, b) => b[1].amount - a[1].amount)
        .slice(0, 10);

    barChart.data.labels = sortedCreditors.map(item => item[0]);
    // Save metadata in the dataset for tooltip access
    barChart.data.datasets[0].data = sortedCreditors.map(item => item[1].amount);
    barChart.data.datasets[0].statusData = sortedCreditors.map(item => Array.from(item[1].statuses).join(', '));
    barChart.update();
};

// ==========================================
// MOCK DATA: For demonstration during setup
// ==========================================
const loadMockData = () => {
    setTimeout(() => {
        allData = [
            { creditor: "สมปอง เซอร์วิส", amount: 15000, status: "ตรงดิว", paymentStatus: "จ่ายแล้ว", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "เจริญ ฮาร์ดแวร์", amount: 8500, status: "ยังไม่ถึงกำหนด", paymentStatus: "รอโอน", category: "รายสัปดาห์", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "การไฟฟ้า", amount: 2300, status: "เกินกำหนด", paymentStatus: "รอโอน", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "A Plus Company", amount: 12293699, status: "ตรงดิว", paymentStatus: "รอโอน", category: "ลิสซิ่ง", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "ค่าเช่าสำนักงาน", amount: 20000, status: "ยังไม่ถึงกำหนด", paymentStatus: "จ่ายแล้ว", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "ผู้รับเหมา กริช", amount: 12000, status: "เกินกำหนด", paymentStatus: "ยกเลิก", category: "รายสัปดาห์", monthDue: "มิ.ย.", yearDue: new Date().getFullYear() },
            { creditor: "สมปอง เซอร์วิส", amount: 7000, status: "ตรงดิว", paymentStatus: "ตัดเช็คผ่าน", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() }
        ];

        loading.classList.add('hidden');
        updateDashboard();
    }, 1000);
};

// Start application
document.addEventListener('DOMContentLoaded', init);

// Export to PDF function (Using browser's native print for perfect Thai font rendering)
const exportToPDF = () => {
    // บันทึกข้อมูลเลขที่เอกสารและวันที่
    const now = new Date();
    const docId = `RT-${now.getTime().toString().slice(-6)}`;
    const dateStr = now.toLocaleString('th-TH');

    // อัปเดตข้อมูลลงในธาตุ HTML สำหรับหน้าพิมพ์
    const printDocId = document.getElementById('printDocId');
    const printIssueDate = document.getElementById('printIssueDate');

    if (printDocId) printDocId.innerText = docId;
    if (printIssueDate) printIssueDate.innerText = dateStr;

    window.print();
};
