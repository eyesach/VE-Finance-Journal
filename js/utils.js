/**
 * Utility functions for the Accounting Journal Calculator
 */

const Utils = {
    /**
     * Format a number as currency (USD)
     * @param {number} amount - The amount to format
     * @returns {string} Formatted currency string
     */
    formatCurrency(amount) {
        return new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD',
            minimumFractionDigits: 2,
            maximumFractionDigits: 2
        }).format(amount || 0);
    },

    /**
     * Format a date as MM/DD/YYYY
     * @param {string} dateStr - Date string (YYYY-MM-DD)
     * @returns {string} Formatted date string
     */
    formatDate(dateStr) {
        if (!dateStr) return '-';
        const date = new Date(dateStr + 'T00:00:00');
        return date.toLocaleDateString('en-US', {
            month: '2-digit',
            day: '2-digit',
            year: 'numeric'
        });
    },

    /**
     * Format a date as short format (MM/DD)
     * @param {string} dateStr - Date string (YYYY-MM-DD)
     * @returns {string} Formatted short date string
     */
    formatDateShort(dateStr) {
        if (!dateStr) return '-';
        const date = new Date(dateStr + 'T00:00:00');
        return date.toLocaleDateString('en-US', {
            month: 'numeric',
            day: 'numeric'
        });
    },

    /**
     * Format a sale date range for display
     * @param {string} startDate - Start date (YYYY-MM-DD)
     * @param {string} endDate - End date (YYYY-MM-DD)
     * @returns {string} Formatted range like "3/1-3/7" or "3/5" for single day
     */
    formatSaleDateRange(startDate, endDate) {
        if (!startDate) return '';
        const start = new Date(startDate + 'T00:00:00');
        const startStr = `${start.getMonth() + 1}/${start.getDate()}`;
        if (!endDate || startDate === endDate) return startStr;
        const end = new Date(endDate + 'T00:00:00');
        return `${startStr}-${end.getMonth() + 1}/${end.getDate()}`;
    },

    /**
     * Get today's date in YYYY-MM-DD format
     * @returns {string} Today's date
     */
    getTodayDate() {
        // Local date, not UTC: toISOString() returns tomorrow during US evenings,
        // which put "Paid Today" on the last day of a month into the next month.
        const d = new Date();
        return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    },

    /**
     * Add whole months to a Date and return the resulting month as YYYY-MM.
     * Works on year/month only, so day-of-month overflow can't skip a month.
     * @param {Date} date
     * @param {number} n - Months to add (may be negative)
     * @returns {string} YYYY-MM
     */
    addMonthsToDate(date, n) {
        const total = date.getFullYear() * 12 + date.getMonth() + n;
        const y = Math.floor(total / 12);
        const m = total - y * 12 + 1;
        return `${y}-${String(m).padStart(2, '0')}`;
    },

    /**
     * Get current month in YYYY-MM format
     * @returns {string} Current month
     */
    getCurrentMonth() {
        const date = new Date();
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    },

    /**
     * Extract month (YYYY-MM) from a date string
     * @param {string} dateStr - Date string (YYYY-MM-DD)
     * @returns {string} Month in YYYY-MM format
     */
    getMonthFromDate(dateStr) {
        if (!dateStr) return null;
        return dateStr.substring(0, 7);
    },

    /**
     * Format a month string to display format (e.g., "OCTOBER 2024")
     * @param {string} monthStr - Month string (YYYY-MM)
     * @returns {string} Formatted month string
     */
    formatMonthDisplay(monthStr) {
        if (!monthStr) return '';
        const [year, month] = monthStr.split('-');
        const date = new Date(year, parseInt(month) - 1, 1);
        return date.toLocaleDateString('en-US', {
            month: 'long',
            year: 'numeric'
        }).toUpperCase();
    },

    /**
     * Format a month string to short display format (e.g., "Jan 2024")
     * @param {string} monthStr - Month string (YYYY-MM)
     * @returns {string} Formatted month string
     */
    formatMonthShort(monthStr) {
        if (!monthStr) return '';
        const [year, month] = monthStr.split('-');
        const date = new Date(year, parseInt(month) - 1, 1);
        return date.toLocaleDateString('en-US', {
            month: 'short',
            year: 'numeric'
        });
    },

    /**
     * Generate an array of month options (past 12 months + next 12 months)
     * @returns {Array} Array of {value: 'YYYY-MM', label: 'Month Year'} objects
     */
    generateMonthOptions() {
        const options = [];
        const today = new Date();

        for (let i = -12; i <= 12; i++) {
            const date = new Date(today.getFullYear(), today.getMonth() + i, 1);
            const value = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            const label = date.toLocaleDateString('en-US', {
                month: 'long',
                year: 'numeric'
            });
            options.push({ value, label });
        }

        return options;
    },

    /**
     * Generate year options for dropdowns (2 years past to 3 years future)
     * @returns {Array} Array of year numbers
     */
    generateYearOptions() {
        const currentYear = new Date().getFullYear();
        const years = [];
        for (let y = currentYear - 2; y <= currentYear + 3; y++) {
            years.push(y);
        }
        return years;
    },

    /**
     * Generate array of YYYY-MM strings from start to end (inclusive)
     * @param {string} startMonth - Start month (YYYY-MM)
     * @param {string} endMonth - End month (YYYY-MM)
     * @returns {Array} Array of YYYY-MM strings
     */
    generateMonthRange(startMonth, endMonth) {
        const result = [];
        const [sy, sm] = startMonth.split('-').map(Number);
        const [ey, em] = endMonth.split('-').map(Number);
        let y = sy, m = sm;
        while (y < ey || (y === ey && m <= em)) {
            result.push(`${y}-${String(m).padStart(2, '0')}`);
            m++;
            if (m > 12) { m = 1; y++; }
        }
        return result;
    },

    /**
     * Filter months array by timeline boundaries
     * @param {Array} months - Array of YYYY-MM strings
     * @param {string|null} start - Timeline start month or null
     * @param {string|null} end - Timeline end month or null
     * @returns {Array} Filtered months
     */
    filterMonthsByTimeline(months, start, end) {
        return months.filter(m => {
            if (start && m < start) return false;
            if (end && m > end) return false;
            return true;
        });
    },

    /**
     * Get years that overlap a timeline range
     * @param {string|null} start - Timeline start (YYYY-MM) or null
     * @param {string|null} end - Timeline end (YYYY-MM) or null
     * @returns {Array} Array of year numbers
     */
    getYearsInTimeline(start, end) {
        if (!start && !end) return this.generateYearOptions();
        const startYear = start ? parseInt(start.split('-')[0]) : new Date().getFullYear() - 2;
        const endYear = end ? parseInt(end.split('-')[0]) : new Date().getFullYear() + 3;
        const years = [];
        for (let y = startYear; y <= endYear; y++) {
            years.push(y);
        }
        return years;
    },

    /**
     * Check if a month string is in the future (after current month)
     * @param {string} month - YYYY-MM string
     * @returns {boolean}
     */
    isFutureMonth(month) {
        return month > this.getCurrentMonth();
    },

    /**
     * Get the next month after the given month
     * @param {string} month - YYYY-MM string
     * @returns {string} Next month YYYY-MM string
     */
    nextMonth(month) {
        const [y, m] = month.split('-').map(Number);
        if (m === 12) return `${y + 1}-01`;
        return `${y}-${String(m + 1).padStart(2, '0')}`;
    },

    /**
     * Convert YYYY-MM to first day of month for date picker min
     * @param {string} month - YYYY-MM string
     * @returns {string} YYYY-MM-01
     */
    timelineToDateMin(month) {
        return month ? `${month}-01` : '';
    },

    /**
     * Convert YYYY-MM to last day of month for date picker max
     * @param {string} month - YYYY-MM string
     * @returns {string} YYYY-MM-{lastDay}
     */
    timelineToDateMax(month) {
        if (!month) return '';
        const [y, m] = month.split('-').map(Number);
        const lastDay = new Date(y, m, 0).getDate();
        return `${month}-${String(lastDay).padStart(2, '0')}`;
    },

    /**
     * Check if payment was late (paid after due month)
     * @param {string} monthDue - Month due (YYYY-MM)
     * @param {string} monthPaid - Month paid (YYYY-MM)
     * @returns {boolean} True if paid late
     */
    isPaidLate(monthDue, monthPaid) {
        if (!monthDue || !monthPaid) return false;
        return monthPaid > monthDue;
    },

    /**
     * Check if a transaction is overdue
     * @param {string} monthDue - Month due (YYYY-MM)
     * @param {string} status - Transaction status
     * @returns {boolean} True if overdue
     */
    isOverdue(monthDue, status) {
        if (!monthDue || status === 'paid' || status === 'received') {
            return false;
        }
        const currentMonth = this.getCurrentMonth();
        return monthDue < currentMonth;
    },

    /**
     * Parse a numeric string to a float
     * @param {string|number} value - Value to parse
     * @returns {number} Parsed float
     */
    parseAmount(value) {
        const parsed = parseFloat(value);
        return isNaN(parsed) ? 0 : parsed;
    },

    /**
     * Generate a unique ID for temporary use
     * @returns {string} Unique ID
     */
    generateTempId() {
        return 'temp_' + Date.now() + '_' + Math.random().toString(36).substring(2, 11);
    },

    /**
     * Debounce a function
     * @param {Function} func - Function to debounce
     * @param {number} wait - Wait time in milliseconds
     * @returns {Function} Debounced function
     */
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    },

    /**
     * Sort transactions by date (newest first)
     * @param {Array} transactions - Array of transactions
     * @returns {Array} Sorted transactions
     */
    sortByDate(transactions) {
        return [...transactions].sort((a, b) => {
            const dateA = a.entry_date || '';
            const dateB = b.entry_date || '';
            return dateB.localeCompare(dateA);
        });
    },

    /**
     * Group transactions by month
     * @param {Array} transactions - Array of transactions
     * @returns {Object} Transactions grouped by month
     */
    groupByMonth(transactions) {
        const groups = {};

        transactions.forEach(transaction => {
            const month = this.getMonthFromDate(transaction.entry_date) || 'Unknown';
            if (!groups[month]) {
                groups[month] = [];
            }
            groups[month].push(transaction);
        });

        const sortedGroups = {};
        Object.keys(groups)
            .sort((a, b) => b.localeCompare(a))
            .forEach(key => {
                sortedGroups[key] = groups[key];
            });

        return sortedGroups;
    },

    /**
     * Group transactions by month due
     * @param {Array} transactions - Array of transactions
     * @returns {Object} Transactions grouped by month due
     */
    groupByMonthDue(transactions) {
        const groups = {};

        transactions.forEach(transaction => {
            const month = transaction.month_due || 'No Due Date';
            if (!groups[month]) {
                groups[month] = [];
            }
            groups[month].push(transaction);
        });

        const sortedGroups = {};
        Object.keys(groups)
            .sort((a, b) => {
                if (a === 'No Due Date') return 1;
                if (b === 'No Due Date') return -1;
                return b.localeCompare(a);
            })
            .forEach(key => {
                sortedGroups[key] = groups[key];
            });

        return sortedGroups;
    },

    /**
     * Group transactions by category
     * @param {Array} transactions - Array of transactions
     * @returns {Object} Transactions grouped by category name
     */
    groupByCategory(transactions) {
        const groups = {};

        transactions.forEach(transaction => {
            const category = transaction.category_name || 'Uncategorized';
            if (!groups[category]) {
                groups[category] = [];
            }
            groups[category].push(transaction);
        });

        const sortedGroups = {};
        Object.keys(groups)
            .sort((a, b) => a.localeCompare(b))
            .forEach(key => {
                sortedGroups[key] = groups[key];
            });

        return sortedGroups;
    },

    /**
     * Get unique months from transactions
     * @param {Array} transactions - Array of transactions
     * @returns {Array} Array of unique months (YYYY-MM format)
     */
    getUniqueMonths(transactions) {
        const months = new Set();
        transactions.forEach(t => {
            const month = this.getMonthFromDate(t.entry_date);
            if (month) months.add(month);
        });
        return Array.from(months).sort((a, b) => b.localeCompare(a));
    },

    /**
     * Escape HTML to prevent XSS
     * @param {string} str - String to escape
     * @returns {string} Escaped string
     */
    escapeHtml(str) {
        if (!str) return '';
        // Escape quotes too: callers interpolate the result into attribute
        // values (value="…", data-notes="…"), where an unescaped " truncates.
        return String(str)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
    },

    /**
     * Sanitize a string for use as a filename
     * @param {string} str - String to sanitize
     * @returns {string} Sanitized string
     */
    sanitizeFilename(str) {
        if (!str) return '';
        return str.replace(/[^a-zA-Z0-9_\- ]/g, '').replace(/\s+/g, '_');
    },

    // ==================== COLOR HELPERS ====================

    /**
     * Convert hex color to HSL
     * @param {string} hex - Hex color (e.g., '#4a90a4')
     * @returns {{h: number, s: number, l: number}} HSL values (h: 0-360, s: 0-100, l: 0-100)
     */
    hexToHSL(hex) {
        hex = hex.replace('#', '');
        const r = parseInt(hex.substring(0, 2), 16) / 255;
        const g = parseInt(hex.substring(2, 4), 16) / 255;
        const b = parseInt(hex.substring(4, 6), 16) / 255;

        const max = Math.max(r, g, b);
        const min = Math.min(r, g, b);
        let h, s;
        const l = (max + min) / 2;

        if (max === min) {
            h = s = 0;
        } else {
            const d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            switch (max) {
                case r: h = ((g - b) / d + (g < b ? 6 : 0)) / 6; break;
                case g: h = ((b - r) / d + 2) / 6; break;
                case b: h = ((r - g) / d + 4) / 6; break;
            }
        }

        return { h: Math.round(h * 360), s: Math.round(s * 100), l: Math.round(l * 100) };
    },

    /**
     * Convert HSL values to hex color
     * @param {number} h - Hue (0-360)
     * @param {number} s - Saturation (0-100)
     * @param {number} l - Lightness (0-100)
     * @returns {string} Hex color string
     */
    hslToHex(h, s, l) {
        s /= 100;
        l /= 100;

        const a = s * Math.min(l, 1 - l);
        const f = (n) => {
            const k = (n + h / 30) % 12;
            const color = l - a * Math.max(Math.min(k - 3, 9 - k, 1), -1);
            return Math.round(255 * color).toString(16).padStart(2, '0');
        };

        return `#${f(0)}${f(8)}${f(4)}`;
    },

    /**
     * Adjust lightness of a hex color
     * @param {string} hex - Hex color
     * @param {number} amount - Lightness adjustment (-100 to 100)
     * @returns {string} Adjusted hex color
     */
    adjustLightness(hex, amount) {
        const hsl = this.hexToHSL(hex);
        hsl.l = Math.max(0, Math.min(100, hsl.l + amount));
        return this.hslToHex(hsl.h, hsl.s, hsl.l);
    },

    // ==================== AMORTIZATION & DEPRECIATION ====================

    /**
     * Compute a full amortization schedule from loan config.
     * Accepts { principal, annual_rate, term_months, payments_per_year, start_date, first_payment_date }.
     * @param {Object} config - Loan configuration
     * @param {Set} [skippedPayments] - Set of payment numbers to skip
     * @param {Object} [paymentOverrides] - Map of paymentNumber => override_amount
     * @returns {Array} Array of payment objects
     */
    computeAmortizationSchedule(config, skippedPayments, paymentOverrides) {
        const { principal, annual_rate, payments_per_year, start_date } = config;
        const termMonths = config.term_months || (config.term_years ? config.term_years * 12 : 0);
        const termYears = termMonths / 12;
        const totalPayments = termYears * payments_per_year;
        const periodicRate = (annual_rate / 100) / payments_per_year;

        const round2 = (v) => Math.round(v * 100) / 100;

        let payment;
        if (periodicRate === 0) {
            payment = round2(principal / totalPayments);
        } else {
            payment = round2(principal * (periodicRate * Math.pow(1 + periodicRate, totalPayments)) /
                      (Math.pow(1 + periodicRate, totalPayments) - 1));
        }

        const schedule = [];
        let balance = round2(principal);

        // Use first_payment_date if provided, otherwise default to start_date + 1 period
        const baseDate = config.first_payment_date
            ? new Date(config.first_payment_date + 'T00:00:00')
            : null;
        const startDate = new Date(start_date + 'T00:00:00');
        const monthsBetween = 12 / payments_per_year;

        for (let i = 1; i <= totalPayments; i++) {
            const interest = round2(balance * periodicRate);
            const isSkipped = skippedPayments && skippedPayments.has(i);

            // Month arithmetic on (year, month) only. Date.setMonth() overflows on the
            // 29th-31st (Jan 31 + 1 month = Mar 3), which skipped February and put two
            // payments in March.
            const month = baseDate
                ? this.addMonthsToDate(baseDate, Math.round(monthsBetween * (i - 1)))
                : this.addMonthsToDate(startDate, Math.round(monthsBetween * i));

            if (isSkipped) {
                // Skipped: interest accrues but no payment made
                balance = round2(balance + interest);
                schedule.push({
                    number: i,
                    month,
                    payment: 0,
                    principal: 0,
                    interest: interest,
                    ending_balance: balance,
                    skipped: true
                });
            } else {
                // Check for payment override
                const hasOverride = paymentOverrides && (i in paymentOverrides);
                const actualPayment = hasOverride ? paymentOverrides[i] : ((i === totalPayments) ? round2(balance + interest) : payment);

                let principalPart = round2(actualPayment - interest);
                if (principalPart < 0) principalPart = 0;

                balance = round2(balance - principalPart);
                if (balance < 0.01) balance = 0;

                schedule.push({
                    number: i,
                    month,
                    payment: actualPayment,
                    principal: principalPart,
                    interest: interest,
                    ending_balance: balance,
                    skipped: false,
                    overridden: hasOverride
                });
            }
        }

        return schedule;
    },

    /**
     * Compute a depreciation schedule for a fixed asset.
     * Returns { [YYYY-MM]: monthlyDeprAmount } for each month in the asset's life.
     * @param {Object} asset - { purchase_cost, salvage_value, useful_life_months, purchase_date, dep_start_date, depreciation_method, is_depreciable }
     * @returns {Object} Map of month string to depreciation amount
     */
    computeDepreciationSchedule(asset) {
        if (!asset.is_depreciable || asset.depreciation_method === 'none') {
            return {};
        }

        const cost = parseFloat(asset.purchase_cost) || 0;
        const salvage = parseFloat(asset.salvage_value) || 0;
        const lifeMonths = parseInt(asset.useful_life_months) || 0;
        if (lifeMonths <= 0 || cost <= salvage) return {};

        const startStr = asset.dep_start_date || asset.purchase_date;
        if (!startStr) return {};

        const startDate = new Date(startStr + 'T00:00:00');
        // Depreciation begins in the month of the start date

        const schedule = {};

        const round2 = (v) => Math.round(v * 100) / 100;

        if (asset.depreciation_method === 'double_declining') {
            const annualRate = 2 / (lifeMonths / 12);
            const monthlyRate = annualRate / 12;
            let bookValue = cost;

            for (let i = 0; i < lifeMonths; i++) {
                const month = this.addMonthsToDate(startDate, i);

                let depr = round2(bookValue * monthlyRate);
                // Don't depreciate below salvage
                if (bookValue - depr < salvage) {
                    depr = round2(bookValue - salvage);
                }
                if (depr < 0.01) break;

                schedule[month] = (schedule[month] || 0) + depr;
                bookValue -= depr;
            }
        } else {
            // straight_line (default)
            const monthlyDepr = round2((cost - salvage) / lifeMonths);

            for (let i = 0; i < lifeMonths; i++) {
                const month = this.addMonthsToDate(startDate, i);
                schedule[month] = (schedule[month] || 0) + monthlyDepr;
            }
        }

        return schedule;
    },

    /**
     * Get a specific month's interest from a pre-computed amortization schedule.
     * @param {Array} schedule - Amortization schedule from computeAmortizationSchedule
     * @param {string} targetMonth - Month in YYYY-MM format
     * @returns {number} Interest for that month (summed if multiple payments in same month)
     */
    getMonthInterestFromSchedule(schedule, targetMonth) {
        return schedule
            .filter(p => p.month === targetMonth)
            .reduce((sum, p) => sum + p.interest, 0);
    },

    /**
     * Convert hex color to RGB string for use in rgba()
     * @param {string} hex - Hex color
     * @returns {string} RGB string (e.g., '74, 144, 164')
     */
    hexToRGBString(hex) {
        hex = hex.replace('#', '');
        const r = parseInt(hex.substring(0, 2), 16);
        const g = parseInt(hex.substring(2, 4), 16);
        const b = parseInt(hex.substring(4, 6), 16);
        return `${r}, ${g}, ${b}`;
    },

    // ==================== BREAK-EVEN ANALYSIS ====================

    /**
     * Core break-even calculation with dual-channel (Consumer + B2B) support.
     * B2B is a committed base load whose contribution offsets fixed costs first;
     * Consumer channel covers the remaining fixed costs.
     * @param {Object} cfg - Break-even config from DB
     * @param {number} monthlyFixedCosts - Total monthly fixed costs
     * @returns {Object} Break-even result
     */
    computeBreakEven(cfg, monthlyFixedCosts) {
        const result = {
            isValid: false,
            consumerCM: 0,
            b2bCM: 0,
            weightedCM: 0,
            consumerCMPercent: 0,
            b2bCMPercent: 0,
            breakEvenUnits: 0,
            breakEvenRevenue: 0,
            monthlyFixedCosts,
            b2bMonthlyRevenue: 0,
            b2bMonthlyCosts: 0,
            b2bMonthlyContribution: 0,
            consumerUnitsNeeded: 0
        };

        const consumer = cfg.consumer || {};
        const b2b = cfg.b2b || {};

        // Consumer contribution margin per unit
        if (consumer.enabled && consumer.avgPrice > 0) {
            result.consumerCM = consumer.avgPrice - (consumer.avgCogs || 0);
            result.consumerCMPercent = (result.consumerCM / consumer.avgPrice) * 100;
        }

        // B2B contribution per month (committed contract)
        if (b2b.enabled && b2b.monthlyUnits > 0 && b2b.ratePerUnit > 0) {
            result.b2bCM = b2b.ratePerUnit - (b2b.cogsPerUnit || 0);
            result.b2bCMPercent = (result.b2bCM / b2b.ratePerUnit) * 100;
            result.b2bMonthlyRevenue = b2b.monthlyUnits * b2b.ratePerUnit;
            result.b2bMonthlyCosts = b2b.monthlyUnits * (b2b.cogsPerUnit || 0);
            result.b2bMonthlyContribution = b2b.monthlyUnits * result.b2bCM;
        }

        // Remaining fixed costs after B2B contribution
        const remainingFixed = Math.max(0, monthlyFixedCosts - result.b2bMonthlyContribution);

        // Consumer units needed to cover remaining fixed costs
        if (consumer.enabled && result.consumerCM > 0) {
            result.consumerUnitsNeededExact = remainingFixed / result.consumerCM;
            result.consumerUnitsNeeded = Math.ceil(result.consumerUnitsNeededExact);
            result.breakEvenUnits = result.consumerUnitsNeeded + (b2b.enabled ? (b2b.monthlyUnits || 0) : 0);
            result.breakEvenRevenue = (result.consumerUnitsNeeded * consumer.avgPrice) + result.b2bMonthlyRevenue;
            result.isValid = true;
        } else if (b2b.enabled && result.b2bMonthlyContribution >= monthlyFixedCosts) {
            // B2B alone covers fixed costs
            result.breakEvenUnits = b2b.monthlyUnits || 0;
            result.breakEvenRevenue = result.b2bMonthlyRevenue;
            result.isValid = true;
        }

        // Weighted CM (blended across both channels at break-even volume)
        if (result.isValid && result.breakEvenUnits > 0) {
            const totalRevenue = result.breakEvenRevenue;
            const totalVarCosts = (result.consumerUnitsNeeded * (consumer.avgCogs || 0)) + result.b2bMonthlyCosts;
            result.weightedCM = totalRevenue > 0 ? ((totalRevenue - totalVarCosts) / totalRevenue) * 100 : 0;
        }

        return result;
    },

    /**
     * Generate data points for the break-even chart and table.
     * B2B units are a constant column (total across timeline); consumer units are the variable X axis.
     * @param {Object} cfg - Break-even config
     * @param {number} monthlyFixedCosts - Average monthly fixed costs
     * @param {number} consumerBEMonthly - Monthly consumer units needed to break even
     * @param {number} unitIncrement - Increment between data points
     * @param {number} monthCount - Number of months in timeline
     * @returns {Array} Array of {consumerUnits, b2bUnits, revenue, variableCosts, fixedCosts, totalCosts}
     */
    computeBreakEvenChartPoints(cfg, monthlyFixedCosts, consumerBEMonthly, unitIncrement, monthCount) {
        const months = Math.max(1, monthCount || 1);
        const increment = Math.max(1, unitIncrement || 100);
        const consumer = cfg.consumer || {};
        const b2b = cfg.b2b || {};

        const b2bMonthly = (b2b.enabled && b2b.monthlyUnits > 0) ? b2b.monthlyUnits : 0;
        const b2bTotal = b2bMonthly * months;
        const fixedTotal = Math.round(monthlyFixedCosts * months * 100) / 100;

        const consumerPrice = consumer.enabled ? (consumer.avgPrice || 0) : 0;
        const consumerCogs = consumer.enabled ? (consumer.avgCogs || 0) : 0;
        const b2bRate = b2b.enabled ? (b2b.ratePerUnit || 0) : 0;
        const b2bCogs = b2b.enabled ? (b2b.cogsPerUnit || 0) : 0;

        const consumerBETotal = Math.ceil(consumerBEMonthly * months);
        const maxConsumer = Math.max(consumerBETotal * 2, increment * 10);
        const points = [];

        for (let cu = 0; cu <= maxConsumer; cu += increment) {
            const revenue = (b2bTotal * b2bRate) + (cu * consumerPrice);
            const variableCosts = (b2bTotal * b2bCogs) + (cu * consumerCogs);
            const totalCosts = variableCosts + fixedTotal;

            points.push({
                consumerUnits: cu,
                b2bUnits: b2bTotal,
                revenue: Math.round(revenue * 100) / 100,
                variableCosts: Math.round(variableCosts * 100) / 100,
                fixedCosts: fixedTotal,
                totalCosts: Math.round(totalCosts * 100) / 100
            });
        }

        return points;
    },

    /**
     * Compute break-even data across a timeline of months.
     * @param {Object} cfg - Break-even config
     * @param {Array} months - Array of YYYY-MM strings
     * @param {Object} assetDeprByMonth - {month: totalDepreciation}
     * @param {Object} loanInterestByMonth - {month: totalInterest}
     * @param {Function} getBudgetForMonth - (month) => budgetFixedCosts
     * @returns {Array} Array of {month, revenue, fixedCosts, variableCosts, totalCosts, profit, breakEvenUnits}
     */
    computeBreakevenTimeline(cfg, months, assetDeprByMonth, loanInterestByMonth, getFixedCostsForMonth) {
        const consumer = cfg.consumer || {};
        const b2b = cfg.b2b || {};
        const results = [];

        months.forEach(month => {
            const fixedCosts = getFixedCostsForMonth(month);

            // Monthly revenue and variable costs from both channels
            const b2bUnits = (b2b.enabled && b2b.monthlyUnits > 0) ? b2b.monthlyUnits : 0;
            const b2bRevenue = b2bUnits * (b2b.ratePerUnit || 0);
            const b2bVarCosts = b2bUnits * (b2b.cogsPerUnit || 0);

            // For timeline, assume break-even consumer volume
            const beResult = this.computeBreakEven(cfg, fixedCosts);
            const consumerUnits = beResult.consumerUnitsNeeded || 0;
            const consumerRevenue = consumerUnits * (consumer.avgPrice || 0);
            const consumerVarCosts = consumerUnits * (consumer.avgCogs || 0);

            const revenue = consumerRevenue + b2bRevenue;
            const variableCosts = consumerVarCosts + b2bVarCosts;
            const totalCosts = fixedCosts + variableCosts;

            results.push({
                month,
                fixedCosts: Math.round(fixedCosts * 100) / 100,
                revenue: Math.round(revenue * 100) / 100,
                variableCosts: Math.round(variableCosts * 100) / 100,
                totalCosts: Math.round(totalCosts * 100) / 100,
                profit: Math.round((revenue - totalCosts) * 100) / 100,
                breakEvenUnits: beResult.breakEvenUnits,
                consumerBERevenue: Math.round(consumerRevenue * 100) / 100
            });
        });

        return results;
    },

    /**
     * Date range covering the previous n full calendar months.
     * @param {number} n - number of months
     * @param {Date} today - reference date
     * @returns {{from:string, to:string}} YYYY-MM-DD bounds
     */
    getLastMonthsRange(n, today) {
        const pad = (x) => String(x).padStart(2, '0');
        const fmt = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
        const start = new Date(today.getFullYear(), today.getMonth() - n, 1);
        const end = new Date(today.getFullYear(), today.getMonth(), 0);
        return { from: fmt(start), to: fmt(end) };
    },

    /**
     * Turn filtered sales lines into weighted average price/COGS, per-product
     * margins and mix, link coverage, unlinked/ignored item groups, and monthly best sellers.
     * Lines: {transaction_no, date, month, name, price, quantity, amount, discount, product_ids, ignored}.
     * A multi-linked line is attributed once, to its lowest known product id.
     * Money is rounded to cents only at output; percentages are 0-100.
     * @param {Array} lines
     * @param {Array<{id:number,name:string,cogs:number}>} products
     * @returns {Object} SalesData (see Sales Data design doc)
     */
    /**
     * Inclusive chronological list of 'YYYY-MM' between two months (date strings accepted).
     * Capped at `max`, keeping the most recent months. Empty if from > to or input invalid.
     */
    monthsBetween(fromYM, toYM, max = 36) {
        const parse = (s) => {
            const m = /^(\d{4})-(\d{2})/.exec(String(s || ''));
            return m ? Number(m[1]) * 12 + (Number(m[2]) - 1) : null;
        };
        const a = parse(fromYM), b = parse(toYM);
        if (a == null || b == null || a > b) return [];
        const start = Math.max(a, b - max + 1);
        const out = [];
        for (let i = start; i <= b; i++) {
            out.push(`${Math.floor(i / 12)}-${String((i % 12) + 1).padStart(2, '0')}`);
        }
        return out;
    },

    computeSalesData(lines, products) {
        const r2 = (x) => Math.round(x * 100) / 100;
        const prodById = new Map((products || []).map(p => [p.id, p]));
        const byProd = new Map();      // id -> accumulator
        const months = new Map();      // month -> Map(id -> {units, rev})
        const unlinkedMap = new Map(), ignoredMap = new Map();
        const linked = { amount: 0, units: 0 }, eligible = { amount: 0, units: 0 };

        const addGroup = (map, l) => {
            const key = l.name + '|' + Number(l.price).toFixed(2);
            let g = map.get(key);
            if (!g) {
                g = { name: l.name, price: l.price, qty: 0, _orders: new Set(), firstDate: l.date, lastDate: l.date };
                map.set(key, g);
            }
            g.qty += l.quantity;
            g._orders.add(l.transaction_no);
            if (l.date < g.firstDate) g.firstDate = l.date;
            if (l.date > g.lastDate) g.lastDate = l.date;
        };
        const finishGroups = (map) => [...map.values()]
            .map(g => ({ name: g.name, price: g.price, qty: g.qty, orders: g._orders.size, firstDate: g.firstDate, lastDate: g.lastDate }))
            .sort((a, b) => b.qty - a.qty);

        (lines || []).forEach(l => {
            if (l.ignored) { addGroup(ignoredMap, l); return; }
            eligible.amount += l.amount;
            eligible.units += l.quantity;
            const known = (l.product_ids || []).filter(id => prodById.has(id)).sort((a, b) => a - b);
            if (!known.length) { addGroup(unlinkedMap, l); return; }
            linked.amount += l.amount;
            linked.units += l.quantity;
            const primary = known[0];
            const rev = l.amount - (l.discount || 0);
            let acc = byProd.get(primary);
            if (!acc) {
                acc = { units: 0, rev: 0, also: new Set(), cogsSet: new Set() };
                byProd.set(primary, acc);
            }
            acc.units += l.quantity;
            acc.rev += rev;
            known.forEach(id => {
                acc.cogsSet.add(Number(prodById.get(id).cogs) || 0);
                if (id !== primary) acc.also.add(prodById.get(id).name);
            });
            if (!months.has(l.month)) months.set(l.month, new Map());
            const mm = months.get(l.month);
            const ma = mm.get(primary) || { units: 0, rev: 0 };
            ma.units += l.quantity;
            ma.rev += rev;
            mm.set(primary, ma);
        });

        let totalUnits = 0, totalRev = 0, totalCogs = 0;
        const rows = [...byProd.entries()].map(([id, a]) => {
            const p = prodById.get(id);
            const cogs = Number(p.cogs) || 0;
            const avgPrice = a.units ? a.rev / a.units : 0;
            const margin = avgPrice - cogs;
            totalUnits += a.units;
            totalRev += a.rev;
            totalCogs += a.units * cogs;
            const cogsVals = [...a.cogsSet].sort((x, y) => x - y);
            return {
                productId: id, name: p.name, alsoNames: [...a.also],
                units: a.units, netRevenue: r2(a.rev), avgPrice: r2(avgPrice), cogs,
                margin: r2(margin), marginPct: avgPrice ? margin / avgPrice * 100 : 0,
                profit: r2(a.units * margin), mix: 0, missingCogs: cogs === 0,
                cogsConflict: cogsVals.length > 1 ? cogsVals : null
            };
        });
        rows.forEach(row => { row.mix = totalUnits ? row.units / totalUnits * 100 : 0; });
        rows.sort((a, b) => b.units - a.units);

        let weighted = null;
        if (totalUnits > 0) {
            const avgPrice = totalRev / totalUnits;
            const avgCogs = totalCogs / totalUnits;
            const cm = avgPrice - avgCogs;
            weighted = {
                units: totalUnits, avgPrice: r2(avgPrice), avgCogs: r2(avgCogs), cm: r2(cm),
                cmPct: avgPrice ? cm / avgPrice * 100 : 0, productCount: rows.length
            };
        }

        const top3 = (arr, key) => arr.slice().sort((a, b) => b[key] - a[key] || a.name.localeCompare(b.name)).slice(0, 3);
        const bestSellers = {};
        [...months.keys()].sort().forEach(m => {
            const items = [...months.get(m).entries()].map(([id, a]) => {
                const p = prodById.get(id);
                return { productId: id, name: p.name, units: a.units, revenue: a.rev, margin: a.rev - a.units * (Number(p.cogs) || 0) };
            });
            const rank = (key, round) => top3(items, key).map(i => ({ productId: i.productId, name: i.name, value: round ? r2(i[key]) : i[key] }));
            bestSellers[m] = { units: rank('units', false), revenue: rank('revenue', true), margin: rank('margin', true) };
        });

        return {
            weighted, byProduct: rows,
            linked: { amount: r2(linked.amount), units: linked.units },
            eligible: { amount: r2(eligible.amount), units: eligible.units },
            linkedPct: eligible.amount > 0 ? linked.amount / eligible.amount * 100 : null,
            unlinked: finishGroups(unlinkedMap), ignored: finishGroups(ignoredMap),
            bestSellers
        };
    },

    /**
     * Evaluate a math expression string with standard operator precedence.
     * Supports: +, -, *, ×, /, ÷, parentheses, unary minus, decimals.
     * Returns { value: number, error: string|null }
     */
    evaluateExpression(str) {
        if (!str || !str.trim()) return { value: 0, error: null };

        // Normalize operator symbols
        let expr = str.replace(/×/g, '*').replace(/÷/g, '/').replace(/−/g, '-');
        // Strip currency symbols and commas from numbers
        expr = expr.replace(/\$\s*/g, '').replace(/,/g, '');

        let pos = 0;
        const ch = () => expr[pos] || '';
        const skipSpaces = () => { while (pos < expr.length && expr[pos] === ' ') pos++; };

        const parseNumber = () => {
            skipSpaces();
            let start = pos;
            if (ch() === '-') pos++;
            if (pos < expr.length && (expr[pos] >= '0' && expr[pos] <= '9' || expr[pos] === '.')) {
                while (pos < expr.length && ((expr[pos] >= '0' && expr[pos] <= '9') || expr[pos] === '.')) pos++;
                return parseFloat(expr.slice(start, pos));
            }
            pos = start;
            return null;
        };

        const parseFactor = () => {
            skipSpaces();
            if (ch() === '(') {
                pos++; // skip '('
                const val = parseExpression();
                skipSpaces();
                if (ch() !== ')') return { value: NaN, error: 'Unmatched parentheses' };
                pos++; // skip ')'
                return val;
            }
            if (ch() === '-') {
                pos++;
                const factor = parseFactor();
                if (factor.error) return factor;
                return { value: -factor.value, error: null };
            }
            const num = parseNumber();
            if (num === null || isNaN(num)) return { value: NaN, error: 'Unexpected character: ' + (ch() || 'end of expression') };
            return { value: num, error: null };
        };

        const parseTerm = () => {
            let result = parseFactor();
            if (result.error) return result;
            skipSpaces();
            while (ch() === '*' || ch() === '/') {
                const op = ch();
                pos++;
                const right = parseFactor();
                if (right.error) return right;
                if (op === '/') {
                    if (right.value === 0) return { value: NaN, error: 'Division by zero' };
                    result.value /= right.value;
                } else {
                    result.value *= right.value;
                }
                skipSpaces();
            }
            return result;
        };

        const parseExpression = () => {
            let result = parseTerm();
            if (result.error) return result;
            skipSpaces();
            while (ch() === '+' || ch() === '-') {
                const op = ch();
                pos++;
                const right = parseTerm();
                if (right.error) return right;
                result.value = op === '+' ? result.value + right.value : result.value - right.value;
                skipSpaces();
            }
            return result;
        };

        try {
            const result = parseExpression();
            skipSpaces();
            if (pos < expr.length) return { value: NaN, error: 'Unexpected character: ' + ch() };
            if (result.error) return result;
            return { value: Math.round(result.value * 100) / 100, error: null };
        } catch (e) {
            return { value: NaN, error: 'Parse error' };
        }
    },

    /**
     * Render minimal markdown (bold, italic, bullets) safely.
     * Order: escape HTML first, THEN transform. Prevents <script> injection.
     * @param {string} body Raw user input
     * @returns {string} HTML string safe for innerHTML
     */
    renderMarkdownSafe(body) {
        if (!body) return '';
        // 1. Escape HTML
        let s = String(body)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#39;');
        // 2. Bold (must run before italic since ** contains *)
        s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
        // 3. Italic
        s = s.replace(/\*([^*]+)\*/g, '<em>$1</em>');
        // 4. Lists: contiguous lines starting with "- " become a <ul>
        const lines = s.split('\n');
        const out = [];
        let inList = false;
        for (const line of lines) {
            if (line.startsWith('- ')) {
                if (!inList) { out.push('<ul>'); inList = true; }
                out.push('<li>' + line.slice(2) + '</li>');
            } else {
                if (inList) { out.push('</ul>'); inList = false; }
                out.push(line);
            }
        }
        if (inList) out.push('</ul>');
        return out.join('\n');
    },

    /**
     * Compute variance flag for a value at index in a series of monthly values.
     * Flag is set if |delta vs prior month| OR |delta vs 3-mo avg| meets/exceeds the threshold.
     * @param {Array<number>} values Monthly values in order
     * @param {number} idx Current index in the series
     * @param {number} thresholdPct Threshold percentage (e.g. 20 for 20%)
     * @returns {Object} { flagged, deltaPrior, deltaAvg3, current }
     *   - deltaPrior: fractional change vs values[idx-1] (null if no prior, Infinity if prior was 0 and current isn't)
     *   - deltaAvg3: fractional change vs avg of values[idx-3..idx-1] (null if <3 priors, Infinity if avg was 0)
     *   - flagged: true if either |delta| >= threshold/100
     */
    computeVariance(values, idx, thresholdPct) {
        const out = { flagged: false, deltaPrior: null, deltaAvg3: null, current: values[idx] };
        if (idx < 1) return out;
        const cur = Number(values[idx]) || 0;
        const prior = Number(values[idx - 1]) || 0;

        // delta vs prior
        if (prior !== 0) {
            out.deltaPrior = (cur - prior) / Math.abs(prior);
        } else if (cur !== 0) {
            out.deltaPrior = Infinity;
        }

        // delta vs 3-mo avg (uses prior 3 months — excludes current).
        // Coerce each value the same way as `prior` above: non-finite/missing → 0.
        // Otherwise a single bad cell would NaN-poison the average and silently disable the flag.
        if (idx >= 3) {
            const slice = values.slice(idx - 3, idx).map(v => Number(v) || 0);
            const avg = slice.reduce((a, b) => a + b, 0) / 3;
            if (avg !== 0) {
                out.deltaAvg3 = (cur - avg) / Math.abs(avg);
            } else if (cur !== 0) {
                out.deltaAvg3 = Infinity;
            }
        }

        const t = thresholdPct / 100;
        if ((out.deltaPrior !== null && Math.abs(out.deltaPrior) >= t) ||
            (out.deltaAvg3 !== null && Math.abs(out.deltaAvg3) >= t)) {
            out.flagged = true;
        }
        return out;
    },

    // ============== ANOMALY RULES (Phase 4) ==============
    // Each rule takes a context object and returns an anomaly or null.
    // Pure functions — no DB access. Composed and run from App.

    anomalyRunwayDrop(ctx) {
        const cur = ctx.cashRunwayDays && ctx.cashRunwayDays.current;
        const prior = ctx.cashRunwayDays && ctx.cashRunwayDays.prior;
        if (cur == null || prior == null) return null;
        const drop = prior - cur;
        if (drop < 10) return null;
        return {
            signature: 'runway_drop',
            severity: 'critical',
            title: `Cash runway shortened by ${Math.round(drop)} days`,
            description: `From ${Math.round(prior)} days last month to ${Math.round(cur)} days this month.`,
            sourceTab: 'cashflow',
            targetMonth: ctx.snapshotMonth,
            target_id: null,
            explanationStarter: `Why did cash runway drop in ${this.formatMonthShort(ctx.snapshotMonth)}?`
        };
    },

    anomalyCategorySpike(ctx) {
        if (!ctx.categories || !ctx.categories.length) return null;
        let best = null;
        let bestMultiplier = 0;
        for (const cat of ctx.categories) {
            const months = cat.recentMonths || [];
            if (months.length < 7) continue; // need 6 priors + current
            const current = Number(months[months.length - 1]) || 0;
            const priors = months.slice(0, months.length - 1);
            const sum = priors.reduce((a, b) => a + (Number(b) || 0), 0);
            const avg = sum / priors.length;
            if (avg <= 0 || current <= 0) continue;
            const multiplier = current / avg;
            if (multiplier >= 3 && multiplier > bestMultiplier) {
                bestMultiplier = multiplier;
                best = { cat, current, avg, multiplier };
            }
        }
        if (!best) return null;
        const mult = best.multiplier;
        const multLabel = Number.isInteger(mult) ? `${mult}x` : `${mult.toFixed(1)}x`;
        return {
            signature: 'category_spike',
            severity: 'notable',
            title: `${best.cat.name} spend ${multLabel} the 6-month avg`,
            description: `$${Math.round(best.current).toLocaleString()} this month vs $${Math.round(best.avg).toLocaleString()} avg.`,
            sourceTab: 'pnl',
            targetMonth: ctx.snapshotMonth,
            target_id: best.cat.id,
            explanationStarter: `Why was ${best.cat.name} spend so high in ${this.formatMonthShort(ctx.snapshotMonth)}?`
        };
    },

    anomalyArAgingTrend(ctx) {
        const a = ctx.arDaysOutstanding;
        if (!a || a.current == null || a.prior == null || a.twoMonthsAgo == null || a.threeMonthsAgo == null) return null;
        const climbing = a.current > a.prior && a.prior > a.twoMonthsAgo && a.twoMonthsAgo > a.threeMonthsAgo;
        if (!climbing) return null;
        return {
            signature: 'ar_aging_trend',
            severity: 'critical',
            title: 'AR Days Outstanding climbing 3 months in a row',
            description: `${Math.round(a.threeMonthsAgo)} → ${Math.round(a.twoMonthsAgo)} → ${Math.round(a.prior)} → ${Math.round(a.current)} days.`,
            sourceTab: 'dashboard',
            targetMonth: ctx.snapshotMonth,
            target_id: null,
            explanationStarter: 'Why has AR Days Outstanding been climbing?'
        };
    },

    anomalyProfitabilityFlip(ctx) {
        const cur = ctx.netIncome && ctx.netIncome.current;
        const prior = ctx.netIncome && ctx.netIncome.prior;
        if (cur == null || prior == null) return null;
        const turnedPositive = prior < 0 && cur >= 0;
        const turnedNegative = prior >= 0 && cur < 0;
        if (!turnedPositive && !turnedNegative) return null;
        return {
            signature: 'profitability_flip',
            severity: turnedPositive ? 'positive' : 'critical',
            title: turnedPositive ? 'First profitable month' : 'First unprofitable month',
            description: turnedPositive
                ? `Net income flipped to +$${Math.round(cur).toLocaleString()} from −$${Math.abs(Math.round(prior)).toLocaleString()}.`
                : `Net income flipped to −$${Math.abs(Math.round(cur)).toLocaleString()} from +$${Math.round(prior).toLocaleString()}.`,
            sourceTab: 'pnl',
            targetMonth: ctx.snapshotMonth,
            target_id: null,
            explanationStarter: turnedPositive
                ? `What drove the profitable month in ${this.formatMonthShort(ctx.snapshotMonth)}?`
                : `Why did we flip unprofitable in ${this.formatMonthShort(ctx.snapshotMonth)}?`
        };
    },

    anomalyDscrBreach(ctx) {
        const cur = ctx.dscr && ctx.dscr.current;
        const prior = ctx.dscr && ctx.dscr.prior;
        if (cur == null || prior == null) return null;
        if (!(cur < 1.0 && prior >= 1.0)) return null;
        return {
            signature: 'dscr_breach',
            severity: 'critical',
            title: 'DSCR dropped below 1.0 — covenant risk',
            description: `From ${prior.toFixed(2)} last month to ${cur.toFixed(2)} this month.`,
            sourceTab: 'dashboard',
            targetMonth: ctx.snapshotMonth,
            target_id: null,
            explanationStarter: `Why did DSCR drop below 1.0 in ${this.formatMonthShort(ctx.snapshotMonth)}?`
        };
    },

    anomalyRecurringDisappeared(ctx) {
        if (!ctx.categories || !ctx.categories.length) return null;
        for (const cat of ctx.categories) {
            const months = cat.recentMonths || [];
            if (months.length < 7) continue;
            const current = Number(months[months.length - 1]) || 0;
            const priors = months.slice(0, months.length - 1);
            const allPriorsNonZero = priors.every(v => (Number(v) || 0) !== 0);
            if (current === 0 && allPriorsNonZero) {
                return {
                    signature: 'recurring_disappeared',
                    severity: 'notable',
                    title: `${cat.name} (recurring) has no entries this month`,
                    description: `${cat.name} had activity in each of the prior 6 months but is $0 in ${this.formatMonthShort(ctx.snapshotMonth)}.`,
                    sourceTab: 'pnl',
                    targetMonth: ctx.snapshotMonth,
                    target_id: cat.id,
                    explanationStarter: `Why no ${cat.name} activity in ${this.formatMonthShort(ctx.snapshotMonth)}?`
                };
            }
        }
        return null;
    }
};

// Export for use in other modules
window.Utils = Utils;
