class EnhancedFeatures {
    constructor() {
        this.comments = new Map();
        this.conditionalFormats = new Map();
        this.activeFilters = new Map();
        this.frozenPanes = { rows: 0, columns: 0 };
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.setupKeyboardShortcuts();
        this.loadStoredData();
    }

    setupEventListeners() {
        document.querySelectorAll('.dropdown-item').forEach(item => {
            item.addEventListener('click', (e) => {
                const action = e.target.closest('[data-action]')?.dataset.action;
                const format = e.target.closest('[data-format]')?.dataset.format;
                
                if (action) this.handleDataAction(action);
                if (format) this.handleFormatting(format);
            });
        });

        document.getElementById('add-comment-btn')?.addEventListener('click', () => {
            this.showCommentPopup();
        });

        document.getElementById('shortcuts-help-btn')?.addEventListener('click', () => {
            this.showShortcutsModal();
        });

        this.setupModalListeners();
        
        this.setupSheetListeners();
    }

    setupModalListeners() {
        document.querySelectorAll('.modal-close, .comment-close').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.target.closest('.modal-overlay, .comment-popup').classList.add('hidden');
            });
        });

        document.querySelector('.comment-save')?.addEventListener('click', () => {
            this.saveComment();
        });

        document.querySelector('.comment-cancel')?.addEventListener('click', () => {
            document.getElementById('comment-popup').classList.add('hidden');
        });

        document.querySelector('.filter-apply')?.addEventListener('click', () => {
            this.applyFilter();
        });

        document.querySelector('.filter-clear')?.addEventListener('click', () => {
            this.clearFilter();
        });

        document.querySelector('.modal-apply')?.addEventListener('click', () => {
            this.applyConditionalFormatting();
        });

        document.querySelector('.modal-cancel')?.addEventListener('click', () => {
            document.getElementById('formatting-modal').classList.add('hidden');
        });
    }

    setupSheetListeners() {
        document.getElementById('sheet-prev')?.addEventListener('click', () => {
            this.navigateSheets(-1);
        });

        document.getElementById('sheet-next')?.addEventListener('click', () => {
            this.navigateSheets(1);
        });

        document.querySelectorAll('.sheet-action-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                const action = e.target.classList.contains('fa-copy') ? 'duplicate' : 'delete';
                this.handleSheetAction(action);
            });
        });
    }

    handleDataAction(action) {
        switch(action) {
            case 'sort-asc':
                this.sortData(true);
                break;
            case 'sort-desc':
                this.sortData(false);
                break;
            case 'filter':
                this.showFilterModal();
                break;
            case 'clear-filter':
                this.clearAllFilters();
                break;
            case 'freeze-rows':
                this.freezePanes(1, 0);
                break;
            case 'freeze-columns':
                this.freezePanes(0, 1);
                break;
            case 'freeze-clear':
                this.freezePanes(0, 0);
                break;
        }
    }

    handleFormatting(format) {
        switch(format) {
            case 'color-scale':
            case 'data-bars':
            case 'greater-than':
            case 'less-than':
                this.showFormattingModal(format);
                break;
            case 'clear':
                this.clearConditionalFormatting();
                break;
        }
    }

    sortData(ascending = true) {
        const activeCell = document.querySelector('.grid__col--active');
        if (!activeCell) {
            this.showNotification('Please select a cell in the column you want to sort');
            return;
        }

        const row = activeCell.closest('.grid__row');
        const cells = Array.from(row.querySelectorAll('.grid__col'));
        const colIndex = cells.indexOf(activeCell);
        const allRows = Array.from(document.querySelectorAll('.grid__row')).slice(1);
        
        const dataWithRows = allRows.map(row => {
            const cell = row.querySelectorAll('.grid__col')[colIndex];
            return {
                row: row,
                value: cell?.textContent || '',
                numeric: parseFloat(cell?.textContent)
            };
        });

        dataWithRows.sort((a, b) => {
            if (!isNaN(a.numeric) && !isNaN(b.numeric)) {
                return ascending ? a.numeric - b.numeric : b.numeric - a.numeric;
            }
            return ascending ? a.value.localeCompare(b.value) : b.value.localeCompare(a.value);
        });

        const container = document.querySelector('.grid__cells-container');
        dataWithRows.forEach((item) => {
            container.appendChild(item.row);
        });

        this.showNotification(`Sorted ${ascending ? 'A→Z' : 'Z→A'}`);
    }

    showFilterModal() {
        document.getElementById('filter-modal').classList.remove('hidden');
    }

    applyFilter() {
        const column = document.querySelector('.filter-column').value;
        const value = document.querySelector('.filter-value').value;
        
        if (!value) {
            this.showNotification('Please enter a filter value');
            return;
        }

        this.activeFilters.set(column, value);
        this.updateFilterDisplay();
        document.getElementById('filter-modal').classList.add('hidden');
        this.showNotification(`Filter applied to column ${column}`);
    }

    clearFilter() {
        const column = document.querySelector('.filter-column').value;
        this.activeFilters.delete(column);
        this.updateFilterDisplay();
        document.getElementById('filter-modal').classList.add('hidden');
        this.showNotification('Filter cleared');
    }

    clearAllFilters() {
        this.activeFilters.clear();
        this.updateFilterDisplay();
        this.showNotification('All filters cleared');
    }

    updateFilterDisplay() {
        const filterStatus = document.getElementById('filter-status');
        filterStatus.classList.toggle('active', this.activeFilters.size > 0);
        filterStatus.classList.toggle('hidden', this.activeFilters.size === 0);
        
        const allRows = Array.from(document.querySelectorAll('.grid__row')).slice(1);
        
        allRows.forEach(row => {
            let shouldShow = true;
            const cells = row.querySelectorAll('.grid__col');
            
            this.activeFilters.forEach((filterValue, column) => {
                const colIndex = column.charCodeAt(0) - 65;
                const cellValue = cells[colIndex]?.textContent || '';
                if (cellValue.toLowerCase().indexOf(filterValue.toLowerCase()) === -1) {
                    shouldShow = false;
                }
            });
            
            row.style.display = shouldShow ? '' : 'none';
        });
    }

    freezePanes(rows, columns) {
        this.frozenPanes = { rows, columns };
        this.updateFreezeDisplay();
        
        if (rows > 0 || columns > 0) {
            this.showNotification(`Frozen ${rows} row(s) and ${columns} column(s)`);
        } else {
            this.showNotification('Freeze panes cleared');
        }
    }

    updateFreezeDisplay() {
        const vertical = document.getElementById('freeze-pane-vertical');
        const horizontal = document.getElementById('freeze-pane-horizontal');
        const corner = document.getElementById('freeze-pane-corner');
        const freezeStatus = document.getElementById('freeze-status');

        vertical.classList.toggle('hidden', this.frozenPanes.columns === 0);
        horizontal.classList.toggle('hidden', this.frozenPanes.rows === 0);
        corner.classList.toggle('hidden', this.frozenPanes.rows === 0 || this.frozenPanes.columns === 0);
        freezeStatus.classList.toggle('active', this.frozenPanes.rows > 0 || this.frozenPanes.columns > 0);
        freezeStatus.classList.toggle('hidden', this.frozenPanes.rows === 0 && this.frozenPanes.columns === 0);

        if (this.frozenPanes.columns > 0) {
            vertical.style.width = `${this.frozenPanes.columns * 100 + 40}px`;
        }
        if (this.frozenPanes.rows > 0) {
            horizontal.style.height = `${this.frozenPanes.rows * 24 + 30}px`;
        }
    }

    showCommentPopup() {
        const activeCell = document.querySelector('.grid__col--active');
        if (!activeCell) {
            this.showNotification('Please select a cell to add a comment');
            return;
        }

        const popup = document.getElementById('comment-popup');
        const rect = activeCell.getBoundingClientRect();
        
        popup.style.left = `${rect.right + 10}px`;
        popup.style.top = `${rect.top}px`;
        popup.classList.remove('hidden');
        
        document.querySelector('.comment-textarea').value = '';
    }

    saveComment() {
        const activeCell = document.querySelector('.grid__col--active');
        const commentText = document.querySelector('.comment-textarea').value;
        
        if (!activeCell || !commentText.trim()) {
            this.showNotification('Please enter a comment');
            return;
        }

        const cellAddress = this.getCellAddress(activeCell);
        this.comments.set(cellAddress, commentText.trim());
        this.updateCommentIndicators();
        
        document.getElementById('comment-popup').classList.add('hidden');
        this.showNotification('Comment added');
    }

    getCellAddress(cell) {
        const row = cell.closest('.grid__row');
        const rowIndex = Array.from(row.parentNode.children).indexOf(row);
        const colIndex = Array.from(row.children).indexOf(cell);
        const column = String.fromCharCode(65 + colIndex);
        return `${column}${rowIndex + 1}`;
    }

    updateCommentIndicators() {
        document.querySelectorAll('.grid__col').forEach(cell => {
            cell.classList.remove('cell-with-comment');
        });

        this.comments.forEach((comment, address) => {
            const cell = this.findCellByAddress(address);
            if (cell) {
                cell.classList.add('cell-with-comment');
                cell.title = comment;
            }
        });
    }

    findCellByAddress(address) {
        const column = address.charCodeAt(0) - 65;
        const row = parseInt(address.slice(1)) - 1;
        const rows = document.querySelectorAll('.grid__row');
        
        if (rows[row]) {
            const cells = rows[row].querySelectorAll('.grid__col');
            return cells[column];
        }
        return null;
    }

    showFormattingModal(type) {
        document.getElementById('formatting-modal').classList.remove('hidden');
    }

    applyConditionalFormatting() {
        const type = document.querySelector('.rule-type').value;
        const value = document.querySelector('.rule-value').value;
        
        const activeCell = document.querySelector('.grid__col--active');
        if (activeCell) {
            activeCell.classList.add(`cell-format-${type}`);
            this.showNotification('Conditional formatting applied');
        }
        
        document.getElementById('formatting-modal').classList.add('hidden');
    }

    clearConditionalFormatting() {
        document.querySelectorAll('.grid__col').forEach(cell => {
            cell.classList.remove('cell-format-color-scale', 'cell-format-data-bars', 
                                'cell-format-greater', 'cell-format-less');
        });
        this.showNotification('Conditional formatting cleared');
    }

    showShortcutsModal() {
        document.getElementById('shortcuts-modal').classList.remove('hidden');
    }

    navigateSheets(direction) {
        this.showNotification(`Sheet navigation: ${direction > 0 ? 'Next' : 'Previous'}`);
    }

    handleSheetAction(action) {
        if (action === 'delete') {
            if (!confirm('Are you sure you want to delete this sheet?')) return;
        }
        this.showNotification(`Sheet ${action} action`);
    }

    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey || e.metaKey) {
                switch(e.key.toLowerCase()) {
                    case 'z':
                        e.preventDefault();
                        this.showNotification('Undo');
                        break;
                    case 'y':
                        e.preventDefault();
                        this.showNotification('Redo');
                        break;
                    case 'c':
                        e.preventDefault();
                        this.copy();
                        break;
                    case 'v':
                        e.preventDefault();
                        this.paste();
                        break;
                    case 'x':
                        e.preventDefault();
                        this.cut();
                        break;
                    case 'f':
                        e.preventDefault();
                        this.showShortcutsModal();
                        break;
                }
            }
            
            if (e.key === 'F2') {
                e.preventDefault();
                const activeCell = document.querySelector('.grid__col--active');
                if (activeCell) {
                    activeCell.focus();
                    const range = document.createRange();
                    range.selectNodeContents(activeCell);
                    const selection = window.getSelection();
                    selection.removeAllRanges();
                    selection.addRange(range);
                }
            }
        });
    }

    copy() {
        const activeCell = document.querySelector('.grid__col--active');
        if (activeCell) {
            navigator.clipboard.writeText(activeCell.textContent).then(() => {
                this.showNotification('Copied to clipboard');
            });
        }
    }

    paste() {
        this.showNotification('Paste ready - click a cell to paste');
    }

    cut() {
        const activeCell = document.querySelector('.grid__col--active');
        if (activeCell) {
            navigator.clipboard.writeText(activeCell.textContent).then(() => {
                activeCell.textContent = '';
                this.showNotification('Cut to clipboard');
            });
        }
    }

    showNotification(message) {
        const notification = document.createElement('div');
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            background: var(--primary-blue);
            color: white;
            padding: 12px 16px;
            border-radius: 4px;
            z-index: 10000;
            font-size: 14px;
            box-shadow: var(--shadow-medium);
        `;
        notification.textContent = message;
        document.body.appendChild(notification);
        
        setTimeout(() => {
            document.body.removeChild(notification);
        }, 3000);
    }

    loadStoredData() {
        try {
            const stored = localStorage.getItem('sheets-enhanced-data');
            if (stored) {
                const data = JSON.parse(stored);
                this.comments = new Map(data.comments || []);
                this.conditionalFormats = new Map(data.formats || []);
                this.activeFilters = new Map(data.filters || []);
                this.frozenPanes = data.frozenPanes || { rows: 0, columns: 0 };
                
                this.updateCommentIndicators();
                this.updateFilterDisplay();
                this.updateFreezeDisplay();
            }
        } catch (e) {
            console.log('No enhanced data found');
        }
    }

    saveStoredData() {
        const data = {
            comments: Array.from(this.comments.entries()),
            formats: Array.from(this.conditionalFormats.entries()),
            filters: Array.from(this.activeFilters.entries()),
            frozenPanes: this.frozenPanes
        };
        localStorage.setItem('sheets-enhanced-data', JSON.stringify(data));
    }
}

document.addEventListener('DOMContentLoaded', () => {
    window.sheetsEnhanced = new EnhancedFeatures();
});

document.addEventListener('click', (e) => {
    if (e.target.classList.contains('modal-overlay')) {
        e.target.classList.add('hidden');
    }
});