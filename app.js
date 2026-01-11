class ScheduleApp {
    constructor() {
        this.currentWeek = this.getCurrentWeek();
        this.currentDay = new Date().getDay()-1;
        if (this.currentDay < 0) this.currentDay = 0;
        this.scheduleData = null;
        this.theme = localStorage.getItem('theme') || 'dark';
        
        this.init();
    }

    init() {
        this.bindEvents();
        this.setTheme(this.theme);
        this.loadSchedule();
        this.renderDays();
    }

    bindEvents() {
        document.getElementById('themeToggle').addEventListener('click', () => {
            this.toggleTheme();
        });

        document.getElementById('refreshBtn').addEventListener('click', () => {
            this.loadSchedule();
        });

        document.getElementById('retryBtn').addEventListener('click', () => {
            this.loadSchedule();
        });

        document.getElementById('weekDropdown').addEventListener('change', (e) => {
            this.currentWeek = parseInt(e.target.value);
            this.renderWeekInfo();
            this.renderSchedule();
        });
    }

    async loadSchedule() {
        this.showLoading();
        
        try {
            const response = await fetch('https://docs.google.com/spreadsheets/d/1NzZJ-EsIW_3A_89i6zjvXrVJuMaGllQZ/export?format=xlsx&gid=2079030459');
            
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            
            const arrayBuffer = await response.arrayBuffer();
            this.scheduleData = await this.parseExcel(arrayBuffer);
            
            this.renderWeeks();
            this.renderWeekInfo();
            this.renderSchedule();
            this.hideError();
            
        } catch (error) {
            console.error('Error loading schedule:', error);
            this.showError();
        }
    }

    async parseExcel(arrayBuffer) {
        try {
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            return this.processWorkbook(workbook);
        } catch (error) {
            console.error('Error parsing Excel:', error);
            return this.getDemoData();
        }
    }

    processWorkbook(workbook) {
        const weeks = {};
        workbook.SheetNames.forEach(sheetName => {  
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            const sheetWeeks = this.parseSheetData(data, sheetName);
            Object.assign(weeks, sheetWeeks);
        });
        return weeks;
    }

    parseSheetData(data, sheetName) {
        const weeks = {};
        const groupColumns = this.findGroupColumns(data);
        
        if (groupColumns.length === 0) {
            return weeks;
        }
        groupColumns.forEach(colIndex => {
            const columnWeeks = this.parseGroupColumn(data, colIndex);
            Object.assign(weeks, columnWeeks);
        });
        
        return weeks;
    }

    findGroupColumns(data) {
        const columns = [];
        if (data.length <= 0) return columns;
        const groupRow = data[0];
        if (!groupRow) return columns;
        
        let foundStart = false;
        
        for (let col = 0; col < groupRow.length; col++) {
            const cellValue = String(groupRow[col] || '').trim();
            if (cellValue.includes('4422')) {
                foundStart = true;
            }
            
            if (foundStart) {
                columns.push(col);
            }
        }
        return columns;
    }

    parseGroupColumn(data, groupCol) {
        const weeks = {};
        const weekNumbers = this.findWeekNumbers(data, groupCol);
        Object.entries(weekNumbers).forEach(([weekNum, weekRow]) => {
            const daySchedules = this.parseWeekSchedule(data, groupCol, weekRow);
            
            if (daySchedules.length > 0) {
                weeks[weekNum] = daySchedules;
            }
        });
        
        return weeks;
    }

    findWeekNumbers(data, groupCol) {
        const weekNumbers = {};
        for (let row = 0; row <= 18; row++) {
            if (row >= data.length) break;
            
            const rowData = data[row];
            if (!rowData || groupCol >= rowData.length) continue;
            
            const cellValue = String(rowData[groupCol] || '').trim();
            
            if (this.isWeekNumber(cellValue)) {
                const weekNum = parseInt(cellValue.replace('.0', ''));
                if (weekNum > 0) {
                    weekNumbers[weekNum] = row;
                }
            }
        }
        
        return weekNumbers;
    }

    isWeekNumber(value) {
        if (!value) return false;
        const cleaned = value.replace('.0', '').trim();
        const num = parseInt(cleaned);
        return !isNaN(num) && num >= 1 && num <= 52;
    }

    parseWeekSchedule(data, groupCol, weekRow) {
        const daySchedules = [];
        let currentDay = '';
        const dayMap = {};
        for (let row = weekRow + 1; row < data.length; row++) {
            if (row >= data.length) break;
            
            const rowData = data[row];
            if (!rowData) continue;
            
            try {
                const dayCell = rowData[0] ? String(rowData[0]).trim() : '';
                
                if (dayCell && this.isDayName(dayCell)) {
                    currentDay = this.normalizeDayName(dayCell);
                    if (!dayMap[currentDay]) {
                        dayMap[currentDay] = [];
                    }
                }
                
                if (!currentDay) continue;
                let timeCell = rowData[1] ? String(rowData[1]).trim() : '';
                switch(timeCell){
                    case "1":
                        timeCell = "8:00-9:30";
                        break;
                    case "2":
                        timeCell = "9:40-11:10";
                        break;
                    case "3":
                        timeCell = "12:00-13:30";
                        break;
                    case "4":
                        timeCell = "13:40-15:10";
                        break;
                    case "5":
                        timeCell = "15:50-17:20";
                        break;
                    case "6":
                        timeCell = "17:30-19:00"
                        break;
                }
                const subjectCell = rowData[groupCol] ? String(rowData[groupCol]).trim() : '';
                
                if (subjectCell && subjectCell !== '-' && subjectCell !== '') {
                    if (subjectCell.toLowerCase().includes('выходной')) {
                        dayMap[currentDay].push({
                            time: '',
                            subject: 'Выходной',
                            teacher: '',
                            classroom: '',
                            isDayOff: true
                        });
                    } else {
                        const parsedInfo = this.parseSubjectInfo(subjectCell);
                        dayMap[currentDay].push({
                            time: timeCell,
                            subject: parsedInfo.subject,
                            teacher: parsedInfo.teacher,
                            classroom: parsedInfo.classroom,
                            isDayOff: false
                        });
                    }
                }
                
            } catch (error) {
                console.warn(`Error parsing row ${row}:`, error);
                continue;
            }
        }
        Object.entries(dayMap).forEach(([dayName, scheduleItems]) => {
            if (scheduleItems.length > 0) {
                daySchedules.push({
                    day: dayName,
                    schedule: scheduleItems
                });
            }
        });
        daySchedules.sort((a, b) => this.getDayOrder(a.day) - this.getDayOrder(b.day));
        
        return daySchedules;
    }

    isDayName(value) {
        if (!value) return false;
        const lower = value.toLowerCase();
        return lower.includes('понедельник') || lower.includes('вторник') || 
               lower.includes('среда') || lower.includes('четверг') ||
               lower.includes('пятница') || lower.includes('суббота') ||
               lower.includes('воскресенье') || 
               ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'].includes(lower);
    }

    normalizeDayName(day) {
        const lower = day.toLowerCase();
        if (lower.includes('понедельник')) return 'ПН';
        if (lower.includes('вторник')) return 'ВТ';
        if (lower.includes('среда')) return 'СР';
        if (lower.includes('четверг')) return 'ЧТ';
        if (lower.includes('пятница')) return 'ПТ';
        if (lower.includes('суббота')) return 'СБ';
        if (lower.includes('воскресенье')) return 'ВС';
        return day.toUpperCase().replace(/[^А-Я]/g, '');
    }

    normalizeTime(time) {
        return time ? time.replace(/\s+/g, ' ').trim() : '';
    }

    parseSubjectInfo(subjectText) {
        const result = {
            subject: subjectText,
            teacher: '',
            classroom: ''
        };
        
        let cleanedText = subjectText.trim();
        const classroomPatterns = [
            /(\d+-\d+)\s/,
            /(\d+[А-Яа-я]*)\s/,
            /ауд\.?\s*(\S+)/i
        ];
        
        for (const pattern of classroomPatterns) {
            const match = cleanedText.match(pattern);
            if (match) {
                result.classroom = match[1];
                cleanedText = cleanedText.replace(match[0], '').trim();
                break;
            }
        }
        const teacherMatch = cleanedText.match(/([А-Я][а-я]+\s[А-Я]\.[А-Я]\.)/);
        if (teacherMatch) {
            result.teacher = teacherMatch[1];
            cleanedText = cleanedText.replace(teacherMatch[0], '').trim();
        }
        
        result.subject = cleanedText.replace(/\s+/g, ' ').trim();
        
        return result;
    }

    getDayOrder(day) {
        const order = { 'ПН': 1, 'ВТ': 2, 'СР': 3, 'ЧТ': 4, 'ПТ': 5, 'СБ': 6, 'ВС': 7 };
        return order[day] || 8;
    }

    getCurrentWeek() {
        const now = new Date();
        const today = now.getDay();
        if (today === 0) {
            const start = new Date(now.getFullYear()-1, 8, 1);
            const diff = Math.floor((now - start) / (1000 * 60 * 60 * 24));
            return Math.floor(diff / 7) + 2;
        }
        const start = new Date(now.getFullYear(), 8, 1);
        const diff = Math.floor((now - start) / (1000 * 60 * 60 * 24));
        return Math.floor(diff / 7) + 1;
    }

    getDemoData() {
        return {
            8: [
                {
                    day: 'ПН',
                    schedule: [
                        {
                            time: '09:00 - 10:30',
                            subject: 'Математика',
                            teacher: 'Иванов А.Б.',
                            classroom: '101',
                            isDayOff: false
                        }
                    ]
                }
            ]
        };
    }

    renderWeeks() {
        const dropdown = document.getElementById('weekDropdown');
        dropdown.innerHTML = '<option value="">Выберите неделю</option>';

        if (this.scheduleData) {
            const weekNumbers = Object.keys(this.scheduleData).map(Number).sort((a, b) => a - b);
            
            weekNumbers.forEach(weekNum => {
                const option = document.createElement('option');
                option.value = weekNum;
                
                if (weekNum === this.currentWeek) {
                    option.classList.add('current-week-option');
                    option.setAttribute('data-current', 'true');
                    option.textContent = `★ Неделя ${weekNum}`;
                } else {
                    option.textContent = `Неделя ${weekNum}`;
                }
                
                if (weekNum === this.currentWeek) {
                    option.selected = true;
                }
                dropdown.appendChild(option);
            });
        }
    }

    renderWeekInfo() {
        this.renderDays();
    }

    getWeekDates(weekNumber) {
        const currentYear = new Date().getFullYear();
        const startOfYear = new Date(currentYear, 8, 1);
        const weekStart = new Date(startOfYear);
        weekStart.setDate(startOfYear.getDate() + (weekNumber - 1) * 7 - startOfYear.getDay() + 1);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekStart.getDate() + 6);
        
        return {
            start: this.formatDate(weekStart),
            end: this.formatDate(weekEnd)
        };
    }

    formatDate(date) {
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        return `${day}.${month}`;
    }

    renderDays() {
        const daysGrid = document.getElementById('daysGrid');
        const days = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ'];
        
        daysGrid.innerHTML = '';

        days.forEach((day, index) => {
            const dayElement = document.createElement('div');
            dayElement.className = `day-chip ${index === this.currentDay ? 'current' : ''} ${index === this.currentDay ? 'selected' : ''}`;
            dayElement.innerHTML = `
                <div class="day-name">${day}</div>
                <div class="day-date">${this.getDayDate(index)}</div>
            `;
            
            dayElement.addEventListener('click', () => {
                this.selectDay(index);
            });
            
            daysGrid.appendChild(dayElement);
        });
    }

    getDayDate(dayIndex) {
        const weekDates = this.getWeekDates(this.currentWeek);
        const [startDay, startMonth] = weekDates.start.split('.').map(Number);
        const startDate = new Date(new Date().getFullYear(), startMonth - 1, startDay);
        const targetDate = new Date(startDate);
        targetDate.setDate(startDate.getDate() + dayIndex);
        
        return `${targetDate.getDate().toString().padStart(2, '0')}.${(targetDate.getMonth() + 1).toString().padStart(2, '0')}`;
    }

    selectDay(dayIndex) {
        this.currentDay = dayIndex;
        
        document.querySelectorAll('.day-chip').forEach((chip, index) => {
            chip.classList.toggle('selected', index === dayIndex);
        });
        
        this.renderSchedule();
    }

    renderSchedule() {
        const scheduleList = document.getElementById('scheduleList');
        
        if (!this.scheduleData || !this.scheduleData[this.currentWeek]) {
            this.showEmptyState();
            return;
        }

        const weekSchedule = this.scheduleData[this.currentWeek];
        const days = ['ПН', 'ВТ', 'СР', 'ЧТ', 'ПТ', 'СБ'];
        const currentDayName = days[this.currentDay];
        const daySchedule = weekSchedule.find(day => day.day === currentDayName);

        if (!daySchedule || daySchedule.schedule.length === 0) {
            this.showEmptyState();
            return;
        }

        scheduleList.innerHTML = '';

        daySchedule.schedule.forEach(item => {
            const scheduleItem = document.createElement('div');
            scheduleItem.className = 'schedule-item';
            
            if (item.isDayOff) {
                scheduleItem.innerHTML = `
                    <div class="schedule-dayoff">
                        <strong>${item.subject}</strong>
                        <p>Можно отдохнуть</p>
                    </div>
                `;
            } else {
                const timeHtml = item.time ? `<div class="schedule-time">${item.time}</div>` : '';
                const teacherHtml = item.teacher ? `<span>${item.teacher}</span>` : '';
                const classroomHtml = item.classroom ? `<span>Аудитория: ${item.classroom}</span>` : '';
                
                scheduleItem.innerHTML = `
                    ${timeHtml}
                    <div class="schedule-details">
                        <div class="schedule-subject">${item.subject}</div>
                        <div class="schedule-meta">
                            ${item.teacher ? `
                                <div class="meta-item">
                                    <img src="icons/person.svg" alt="Преподаватель" class="meta-icon">
                                    <span>${item.teacher}</span>
                                </div>
                            ` : ''}
                            ${item.classroom ? `
                                <div class="meta-item">
                                    <img src="icons/room.svg" alt="Аудитория" class="meta-icon">
                                    <span>${item.classroom}</span>
                                </div>
                            ` : ''}
                        </div>
                    </div>
                `;
            }
            
            scheduleList.appendChild(scheduleItem);
        });

        this.showSchedule();
    }

    toggleTheme() {
        this.theme = this.theme === 'light' ? 'dark' : 'light';
        this.setTheme(this.theme);
        localStorage.setItem('theme', this.theme);
    }

    setTheme(theme) {
        document.documentElement.setAttribute('data-theme', theme);
        const themeIcon = document.getElementById('themeIcon');
        themeIcon.src = theme === 'light' ? 'icons/dark_mode.svg' : 'icons/light_mode.svg';
    }

    showLoading() {
        document.getElementById('loading').classList.remove('hidden');
        document.getElementById('errorMessage').classList.add('hidden');
        document.getElementById('scheduleList').classList.add('hidden');
        document.getElementById('emptyState').classList.add('hidden');
    }

    showError() {
        document.getElementById('loading').classList.add('hidden');
        document.getElementById('errorMessage').classList.remove('hidden');
        document.getElementById('scheduleList').classList.add('hidden');
        document.getElementById('emptyState').classList.add('hidden');
    }

    showSchedule() {
        document.getElementById('loading').classList.add('hidden');
        document.getElementById('errorMessage').classList.add('hidden');
        document.getElementById('scheduleList').classList.remove('hidden');
        document.getElementById('emptyState').classList.add('hidden');
    }

    showEmptyState() {
        document.getElementById('loading').classList.add('hidden');
        document.getElementById('errorMessage').classList.add('hidden');
        document.getElementById('scheduleList').classList.add('hidden');
        document.getElementById('emptyState').classList.remove('hidden');
    }

    hideError() {
        document.getElementById('errorMessage').classList.add('hidden');
    }
}

document.addEventListener('DOMContentLoaded', () => {
    new ScheduleApp();
});
