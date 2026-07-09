import * as XLSX from 'xlsx'
import { useEffect, useMemo, useState } from 'react'
import type { ChangeEvent, FormEvent } from 'react'
import './App.css'

type Expense = {
  id: string
  date: string
  category: string
  description: string
  amount: number
}

type BackupData = {
  appName: string
  version: string
  createdAt: string
  expenses: Expense[]
  customCategories?: string[]
}

type ImportedData = {
  expenses: Expense[]
  customCategories: string[]
}

type AppSection = 'dashboard' | 'expenses' | 'reports' | 'categories' | 'settings'

const STORAGE_KEY = 'expense-tracker-expenses'
const CATEGORY_STORAGE_KEY = 'expense-tracker-custom-categories'
const FALLBACK_CATEGORY = 'Другое'

const DEFAULT_CATEGORIES = [
  'Продукты',
  'Кафе и рестораны',
  'Транспорт',
  'Дом',
  'Здоровье',
  'Одежда',
  'Развлечения',
  'Подписки',
  'Обучение',
  FALLBACK_CATEGORY,
]

const MONTH_OPTIONS = [
  { value: '01', label: 'Январь' },
  { value: '02', label: 'Февраль' },
  { value: '03', label: 'Март' },
  { value: '04', label: 'Апрель' },
  { value: '05', label: 'Май' },
  { value: '06', label: 'Июнь' },
  { value: '07', label: 'Июль' },
  { value: '08', label: 'Август' },
  { value: '09', label: 'Сентябрь' },
  { value: '10', label: 'Октябрь' },
  { value: '11', label: 'Ноябрь' },
  { value: '12', label: 'Декабрь' },
]

const NAV_ITEMS: Array<{ id: AppSection; label: string }> = [
  { id: 'dashboard', label: 'Главная' },
  { id: 'expenses', label: 'Расходы' },
  { id: 'reports', label: 'Отчеты' },
  { id: 'categories', label: 'Категории' },
  { id: 'settings', label: 'Настройки' },
]

const ALL_CATEGORIES_OPTION = 'Все категории'

function getTodayDate() {
  return new Date().toISOString().slice(0, 10)
}

function getCurrentMonth() {
  return new Date().toISOString().slice(5, 7)
}

function getCurrentYear() {
  return String(new Date().getFullYear())
}

function getCurrentMonthStartDate() {
  return `${getCurrentYear()}-${getCurrentMonth()}-01`
}

function sanitizeAmountInput(value: string) {
  return value.replace(/[^\d.,]/g, '').replace(',', '.')
}

function normalizeCategoryName(value: string) {
  return value.trim().replace(/\s+/g, ' ')
}

function hasCategoryName(categories: string[], name: string) {
  const normalizedName = normalizeCategoryName(name).toLowerCase()

  return categories.some((category) => category.toLowerCase() === normalizedName)
}

function getCleanCustomCategories(candidate: unknown) {
  if (!Array.isArray(candidate)) {
    return []
  }

  const result: string[] = []

  candidate.forEach((item) => {
    if (typeof item !== 'string') {
      return
    }

    const categoryName = normalizeCategoryName(item)

    if (!categoryName) {
      return
    }

    if (hasCategoryName(DEFAULT_CATEGORIES, categoryName)) {
      return
    }

    if (hasCategoryName(result, categoryName)) {
      return
    }

    result.push(categoryName)
  })

  return result
}

function formatCurrency(value: number) {
  return new Intl.NumberFormat('ru-RU', {
    style: 'currency',
    currency: 'RUB',
    maximumFractionDigits: 2,
  }).format(value)
}

function formatDate(date: string) {
  return new Intl.DateTimeFormat('ru-RU', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric',
  }).format(new Date(`${date}T00:00:00`))
}

function formatDateForFile(date: string) {
  const [year, month, day] = date.split('-')

  return `${day}.${month}.${year}`
}

function toExcelDate(date: string) {
  return new Date(`${date}T00:00:00`)
}

function getDaysInclusive(startDate: string, endDate: string) {
  if (!startDate || !endDate || startDate > endDate) {
    return 1
  }

  const start = toExcelDate(startDate).getTime()
  const end = toExcelDate(endDate).getTime()
  const millisecondsInDay = 24 * 60 * 60 * 1000

  return Math.floor((end - start) / millisecondsInDay) + 1
}

function getBackupFileName() {
  const now = new Date()
  const day = String(now.getDate()).padStart(2, '0')
  const month = String(now.getMonth() + 1).padStart(2, '0')
  const year = now.getFullYear()
  const hours = String(now.getHours()).padStart(2, '0')
  const minutes = String(now.getMinutes()).padStart(2, '0')

  return `expense-tracker-backup-${day}-${month}-${year}-${hours}-${minutes}.json`
}

function isExpense(candidate: unknown): candidate is Expense {
  if (typeof candidate !== 'object' || candidate === null) {
    return false
  }

  const expense = candidate as Partial<Expense>

  return (
    typeof expense.id === 'string' &&
    typeof expense.date === 'string' &&
    typeof expense.category === 'string' &&
    typeof expense.description === 'string' &&
    typeof expense.amount === 'number'
  )
}

function getImportedData(candidate: unknown): ImportedData | null {
  if (Array.isArray(candidate) && candidate.every(isExpense)) {
    return {
      expenses: candidate,
      customCategories: [],
    }
  }

  if (typeof candidate !== 'object' || candidate === null) {
    return null
  }

  const backup = candidate as Partial<BackupData>

  if (Array.isArray(backup.expenses) && backup.expenses.every(isExpense)) {
    return {
      expenses: backup.expenses,
      customCategories: getCleanCustomCategories(backup.customCategories),
    }
  }

  return null
}

function setColumnWidths(sheet: XLSX.WorkSheet, widths: number[]) {
  sheet['!cols'] = widths.map((width) => ({ wch: width }))
}

function App() {
  const [activeSection, setActiveSection] = useState<AppSection>('dashboard')

  const [expenses, setExpenses] = useState<Expense[]>(() => {
    const savedExpenses = localStorage.getItem(STORAGE_KEY)

    if (!savedExpenses) {
      return []
    }

    try {
      return JSON.parse(savedExpenses) as Expense[]
    } catch {
      return []
    }
  })

  const [customCategories, setCustomCategories] = useState<string[]>(() => {
    const savedCategories = localStorage.getItem(CATEGORY_STORAGE_KEY)

    if (!savedCategories) {
      return []
    }

    try {
      return getCleanCustomCategories(JSON.parse(savedCategories))
    } catch {
      return []
    }
  })

  const [date, setDate] = useState(getTodayDate())
  const [category, setCategory] = useState(DEFAULT_CATEGORIES[0])
  const [description, setDescription] = useState('')
  const [amount, setAmount] = useState('')
  const [message, setMessage] = useState('')
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonth())
  const [selectedYear, setSelectedYear] = useState(getCurrentYear())
  const [selectedCategory, setSelectedCategory] = useState(ALL_CATEGORIES_OPTION)

  const [editingExpenseId, setEditingExpenseId] = useState<string | null>(null)
  const [editDate, setEditDate] = useState('')
  const [editCategory, setEditCategory] = useState(DEFAULT_CATEGORIES[0])
  const [editDescription, setEditDescription] = useState('')
  const [editAmount, setEditAmount] = useState('')

  const [newCategoryName, setNewCategoryName] = useState('')
  const [editingCategoryName, setEditingCategoryName] = useState<string | null>(null)
  const [editedCategoryName, setEditedCategoryName] = useState('')

  const [reportStartDate, setReportStartDate] = useState(getCurrentMonthStartDate())
  const [reportEndDate, setReportEndDate] = useState(getTodayDate())
  const [reportCategory, setReportCategory] = useState(ALL_CATEGORIES_OPTION)

  const allCategories = useMemo(() => {
    return [...DEFAULT_CATEGORIES, ...customCategories]
  }, [customCategories])

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(expenses))
  }, [expenses])

  useEffect(() => {
    localStorage.setItem(CATEGORY_STORAGE_KEY, JSON.stringify(customCategories))
  }, [customCategories])

  useEffect(() => {
    if (!message) {
      return
    }

    const timerId = window.setTimeout(() => {
      setMessage('')
    }, 2000)

    return () => window.clearTimeout(timerId)
  }, [message])

  const yearOptions = useMemo(() => {
    const currentYear = new Date().getFullYear()
    const yearsFromExpenses = expenses.map((expense) =>
      Number(expense.date.slice(0, 4)),
    )

    const minYear = Math.min(currentYear - 2, ...yearsFromExpenses)
    const maxYear = Math.max(currentYear + 2, ...yearsFromExpenses)

    const years: string[] = []

    for (let year = maxYear; year >= minYear; year -= 1) {
      years.push(String(year))
    }

    return years
  }, [expenses])

  const selectedPeriod = `${selectedYear}-${selectedMonth}`

  const filteredExpenses = useMemo(() => {
    return expenses.filter((expense) => {
      const isSameMonth = expense.date.startsWith(selectedPeriod)
      const isSameCategory =
        selectedCategory === ALL_CATEGORIES_OPTION ||
        expense.category === selectedCategory

      return isSameMonth && isSameCategory
    })
  }, [expenses, selectedPeriod, selectedCategory])

  const totalAmount = useMemo(() => {
    return filteredExpenses.reduce((sum, expense) => sum + expense.amount, 0)
  }, [filteredExpenses])

  const averageAmount =
    filteredExpenses.length > 0 ? totalAmount / filteredExpenses.length : 0

  const sortedExpenses = useMemo(() => {
    return [...filteredExpenses].sort((a, b) => b.date.localeCompare(a.date))
  }, [filteredExpenses])

  const latestExpenses = sortedExpenses.slice(0, 3)

  const isReportDateRangeInvalid = reportStartDate > reportEndDate

  const reportExpenses = useMemo(() => {
    if (isReportDateRangeInvalid) {
      return []
    }

    return expenses
      .filter((expense) => {
        const isInDateRange =
          expense.date >= reportStartDate && expense.date <= reportEndDate
        const isSameCategory =
          reportCategory === ALL_CATEGORIES_OPTION ||
          expense.category === reportCategory

        return isInDateRange && isSameCategory
      })
      .sort((a, b) => a.date.localeCompare(b.date))
  }, [expenses, reportStartDate, reportEndDate, reportCategory, isReportDateRangeInvalid])

  const reportTotalAmount = useMemo(() => {
    return reportExpenses.reduce((sum, expense) => sum + expense.amount, 0)
  }, [reportExpenses])

  const reportDaysCount = getDaysInclusive(reportStartDate, reportEndDate)
  const reportAveragePerDay = reportTotalAmount / reportDaysCount

  const reportCategorySummary = useMemo(() => {
    const summary = new Map<string, { category: string; count: number; amount: number }>()

    reportExpenses.forEach((expense) => {
      const current = summary.get(expense.category) ?? {
        category: expense.category,
        count: 0,
        amount: 0,
      }

      summary.set(expense.category, {
        ...current,
        count: current.count + 1,
        amount: current.amount + expense.amount,
      })
    })

    return Array.from(summary.values())
      .map((item) => ({
        ...item,
        percent: reportTotalAmount > 0 ? (item.amount / reportTotalAmount) * 100 : 0,
      }))
      .sort((a, b) => b.amount - a.amount)
  }, [reportExpenses, reportTotalAmount])

  const reportDailySummary = useMemo(() => {
    const summary = new Map<
      string,
      { date: string; count: number; amount: number; expenses: Expense[] }
    >()

    reportExpenses.forEach((expense) => {
      const current = summary.get(expense.date) ?? {
        date: expense.date,
        count: 0,
        amount: 0,
        expenses: [],
      }

      summary.set(expense.date, {
        ...current,
        count: current.count + 1,
        amount: current.amount + expense.amount,
        expenses: [...current.expenses, expense],
      })
    })

    return Array.from(summary.values()).sort((a, b) => a.date.localeCompare(b.date))
  }, [reportExpenses])

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    const numericAmount = Number(amount.replace(',', '.'))

    if (!date || !category || !numericAmount || numericAmount <= 0) {
      setMessage('Заполните дату, категорию и сумму больше 0')
      return
    }

    const newExpense: Expense = {
      id: crypto.randomUUID(),
      date,
      category,
      description: description.trim() || 'Без описания',
      amount: numericAmount,
    }

    setExpenses((currentExpenses) => [newExpense, ...currentExpenses])

    setDescription('')
    setAmount('')
    setMessage('Расход успешно добавлен')
  }

  function startEditing(expense: Expense) {
    setEditingExpenseId(expense.id)
    setEditDate(expense.date)
    setEditCategory(expense.category)
    setEditDescription(expense.description)
    setEditAmount(String(expense.amount))
  }

  function cancelEditing() {
    setEditingExpenseId(null)
    setEditDate('')
    setEditCategory(DEFAULT_CATEGORIES[0])
    setEditDescription('')
    setEditAmount('')
  }

  function handleEditSubmit(event: FormEvent<HTMLFormElement>, id: string) {
    event.preventDefault()

    const numericAmount = Number(editAmount.replace(',', '.'))

    if (!editDate || !editCategory || !numericAmount || numericAmount <= 0) {
      setMessage('Заполните дату, категорию и сумму больше 0')
      return
    }

    setExpenses((currentExpenses) =>
      currentExpenses.map((expense) =>
        expense.id === id
          ? {
              ...expense,
              date: editDate,
              category: editCategory,
              description: editDescription.trim() || 'Без описания',
              amount: numericAmount,
            }
          : expense,
      ),
    )

    cancelEditing()
    setMessage('Расход обновлен')
  }

  function handleDelete(id: string) {
    const shouldDelete = window.confirm('Удалить этот расход?')

    if (!shouldDelete) {
      return
    }

    setExpenses((currentExpenses) =>
      currentExpenses.filter((expense) => expense.id !== id),
    )

    if (editingExpenseId === id) {
      cancelEditing()
    }

    setMessage('Расход удален')
  }

  function handleAddCategory(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    const categoryName = normalizeCategoryName(newCategoryName)

    if (!categoryName) {
      setMessage('Введите название категории')
      return
    }

    if (hasCategoryName(allCategories, categoryName)) {
      setMessage('Такая категория уже существует')
      return
    }

    setCustomCategories((currentCategories) => [...currentCategories, categoryName])
    setNewCategoryName('')
    setCategory(categoryName)
    setMessage('Категория добавлена')
  }

  function startCategoryEditing(categoryName: string) {
    setEditingCategoryName(categoryName)
    setEditedCategoryName(categoryName)
  }

  function cancelCategoryEditing() {
    setEditingCategoryName(null)
    setEditedCategoryName('')
  }

  function handleRenameCategory(event: FormEvent<HTMLFormElement>) {
    event.preventDefault()

    if (!editingCategoryName) {
      return
    }

    const oldCategoryName = editingCategoryName
    const newName = normalizeCategoryName(editedCategoryName)

    if (!newName) {
      setMessage('Введите новое название категории')
      return
    }

    if (oldCategoryName.toLowerCase() === newName.toLowerCase()) {
      cancelCategoryEditing()
      return
    }

    const categoriesWithoutCurrent = allCategories.filter(
      (currentCategory) => currentCategory !== oldCategoryName,
    )

    if (hasCategoryName(categoriesWithoutCurrent, newName)) {
      setMessage('Такая категория уже существует')
      return
    }

    setCustomCategories((currentCategories) =>
      currentCategories.map((currentCategory) =>
        currentCategory === oldCategoryName ? newName : currentCategory,
      ),
    )

    setExpenses((currentExpenses) =>
      currentExpenses.map((expense) =>
        expense.category === oldCategoryName
          ? {
              ...expense,
              category: newName,
            }
          : expense,
      ),
    )

    if (category === oldCategoryName) {
      setCategory(newName)
    }

    if (selectedCategory === oldCategoryName) {
      setSelectedCategory(newName)
    }

    if (reportCategory === oldCategoryName) {
      setReportCategory(newName)
    }

    if (editCategory === oldCategoryName) {
      setEditCategory(newName)
    }

    cancelCategoryEditing()
    setMessage('Категория переименована')
  }

  function handleDeleteCategory(categoryName: string) {
    if (hasCategoryName(DEFAULT_CATEGORIES, categoryName)) {
      setMessage('Системную категорию нельзя удалить')
      return
    }

    const isUsed = expenses.some((expense) => expense.category === categoryName)

    const shouldDelete = window.confirm(
      isUsed
        ? `Категория «${categoryName}» уже используется в расходах. Расходы будут перенесены в категорию «${FALLBACK_CATEGORY}». Продолжить?`
        : `Удалить категорию «${categoryName}»?`,
    )

    if (!shouldDelete) {
      return
    }

    setCustomCategories((currentCategories) =>
      currentCategories.filter((currentCategory) => currentCategory !== categoryName),
    )

    if (isUsed) {
      setExpenses((currentExpenses) =>
        currentExpenses.map((expense) =>
          expense.category === categoryName
            ? {
                ...expense,
                category: FALLBACK_CATEGORY,
              }
            : expense,
        ),
      )
    }

    if (category === categoryName) {
      setCategory(FALLBACK_CATEGORY)
    }

    if (selectedCategory === categoryName) {
      setSelectedCategory(ALL_CATEGORIES_OPTION)
    }

    if (reportCategory === categoryName) {
      setReportCategory(ALL_CATEGORIES_OPTION)
    }

    if (editCategory === categoryName) {
      setEditCategory(FALLBACK_CATEGORY)
    }

    if (editingCategoryName === categoryName) {
      cancelCategoryEditing()
    }

    setMessage('Категория удалена')
  }

  function handleCreateBackup() {
    const backupData: BackupData = {
      appName: 'Expense Tracker',
      version: '1.3.0',
      createdAt: new Date().toISOString(),
      expenses,
      customCategories,
    }

    const backupContent = JSON.stringify(backupData, null, 2)
    const blob = new Blob([backupContent], { type: 'application/json' })
    const url = URL.createObjectURL(blob)
    const link = document.createElement('a')

    link.href = url
    link.download = getBackupFileName()
    link.click()

    URL.revokeObjectURL(url)

    setMessage('Резервная копия создана')
  }

  async function handleRestoreBackup(event: ChangeEvent<HTMLInputElement>) {
    const input = event.currentTarget
    const file = input.files?.[0]

    if (!file) {
      return
    }

    try {
      const content = await file.text()
      const parsedData = JSON.parse(content) as unknown
      const importedData = getImportedData(parsedData)

      if (!importedData) {
        setMessage('Файл резервной копии некорректен')
        return
      }

      const shouldRestore = window.confirm(
        'Текущие расходы и пользовательские категории будут заменены данными из резервной копии. Продолжить?',
      )

      if (!shouldRestore) {
        return
      }

      setExpenses(importedData.expenses)
      setCustomCategories(importedData.customCategories)
      setSelectedCategory(ALL_CATEGORIES_OPTION)
      setReportCategory(ALL_CATEGORIES_OPTION)
      setCategory(DEFAULT_CATEGORIES[0])
      cancelEditing()
      cancelCategoryEditing()
      setMessage('Данные восстановлены')
    } catch {
      setMessage('Не удалось восстановить данные из файла')
    } finally {
      input.value = ''
    }
  }

  function handleDownloadExcel() {
    if (isReportDateRangeInvalid) {
      setMessage('Дата начала отчета не может быть позже даты окончания')
      return
    }

    if (reportExpenses.length === 0) {
      setMessage('Нет данных для Excel-отчета')
      return
    }

    const workbook = XLSX.utils.book_new()

    const summaryRows = [
      ['Отчет по расходам'],
      [],
      ['Период', `${formatDate(reportStartDate)} — ${formatDate(reportEndDate)}`],
      ['Категория', reportCategory],
      ['Дата формирования', new Date()],
      ['Количество операций', reportExpenses.length],
      ['Количество дней в периоде', reportDaysCount],
      ['Общая сумма расходов', reportTotalAmount],
      ['Средний расход в день', reportAveragePerDay],
      [],
      ['Расходы по категориям'],
      ['Категория', 'Количество операций', 'Сумма', 'Доля от общих расходов (%)'],
      ...reportCategorySummary.map((item) => [
        item.category,
        item.count,
        item.amount,
        Number(item.percent.toFixed(2)),
      ]),
      [],
      ['Расходы по дням'],
      ['Дата', 'Количество операций', 'Сумма'],
      ...reportDailySummary.map((item) => [
        toExcelDate(item.date),
        item.count,
        item.amount,
      ]),
    ]

    const summarySheet = XLSX.utils.aoa_to_sheet(summaryRows)
    setColumnWidths(summarySheet, [28, 24, 18, 28])
    summarySheet['!autofilter'] = { ref: 'A12:D12' }

    const detailRows = [
      ['№', 'Дата', 'Категория', 'Описание', 'Сумма'],
      ...reportExpenses.map((expense, index) => [
        index + 1,
        toExcelDate(expense.date),
        expense.category,
        expense.description,
        expense.amount,
      ]),
      [],
      ['', '', '', 'Итого', reportTotalAmount],
    ]

    const detailSheet = XLSX.utils.aoa_to_sheet(detailRows)
    setColumnWidths(detailSheet, [8, 16, 24, 42, 18])
    detailSheet['!autofilter'] = { ref: 'A1:E1' }

    const dailyRows: Array<Array<string | number | Date>> = [
      ['Дата', 'Описание', 'Категория', 'Сумма'],
    ]

    reportDailySummary.forEach((day) => {
      dailyRows.push([toExcelDate(day.date), 'Итого за день', '', day.amount])

      day.expenses.forEach((expense) => {
        dailyRows.push([
          toExcelDate(expense.date),
          expense.description,
          expense.category,
          expense.amount,
        ])
      })

      dailyRows.push([])
    })

    dailyRows.push(['', '', 'Общая сумма периода', reportTotalAmount])

    const dailySheet = XLSX.utils.aoa_to_sheet(dailyRows)
    setColumnWidths(dailySheet, [16, 42, 24, 18])
    dailySheet['!autofilter'] = { ref: 'A1:D1' }

    XLSX.utils.book_append_sheet(workbook, summarySheet, 'Сводка')
    XLSX.utils.book_append_sheet(workbook, detailSheet, 'Детализация')
    XLSX.utils.book_append_sheet(workbook, dailySheet, 'По дням')

    XLSX.writeFile(
      workbook,
      `Отчет по расходам_${formatDateForFile(reportStartDate)}-${formatDateForFile(
        reportEndDate,
      )}.xlsx`,
      { cellDates: true },
    )

    setMessage('Excel-отчет скачан')
  }

  return (
    <main className="app">
      <section className="app__container">
        <header className="app__header">
          <p className="app__eyebrow">Личный финансовый помощник</p>
          <h1>Учет расходов</h1>
          <p className="app__description">
            Простое приложение для фиксации расходов, анализа трат и формирования отчетов.
          </p>
        </header>

        <nav className="app-nav">
          {NAV_ITEMS.map((item) => (
            <button
              className={`app-nav__button ${
                activeSection === item.id ? 'app-nav__button--active' : ''
              }`}
              key={item.id}
              type="button"
              onClick={() => setActiveSection(item.id)}
            >
              {item.label}
            </button>
          ))}
        </nav>

        {message && (
          <div className="toast">
            <span>{message}</span>
            <button type="button" onClick={() => setMessage('')}>
              ×
            </button>
          </div>
        )}

        {activeSection === 'dashboard' && (
          <>
            <section className="section-heading">
              <h2>Главная</h2>
              <p>Краткая сводка по выбранному месяцу.</p>
            </section>

            <section className="filters">
              <label>
                <span>Месяц</span>
                <select
                  value={selectedMonth}
                  onChange={(event) => setSelectedMonth(event.target.value)}
                >
                  {MONTH_OPTIONS.map((month) => (
                    <option key={month.value} value={month.value}>
                      {month.label}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                <span>Год</span>
                <select
                  value={selectedYear}
                  onChange={(event) => setSelectedYear(event.target.value)}
                >
                  {yearOptions.map((year) => (
                    <option key={year} value={year}>
                      {year}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                <span>Категория</span>
                <select
                  value={selectedCategory}
                  onChange={(event) => setSelectedCategory(event.target.value)}
                >
                  <option value={ALL_CATEGORIES_OPTION}>{ALL_CATEGORIES_OPTION}</option>
                  {allCategories.map((currentCategory) => (
                    <option key={currentCategory} value={currentCategory}>
                      {currentCategory}
                    </option>
                  ))}
                </select>
              </label>
            </section>

            <section className="summary">
              <article className="summary__card">
                <span>Расходы за период</span>
                <strong>{formatCurrency(totalAmount)}</strong>
              </article>

              <article className="summary__card">
                <span>Количество записей</span>
                <strong>{filteredExpenses.length}</strong>
              </article>

              <article className="summary__card">
                <span>Средний расход</span>
                <strong>{formatCurrency(averageAmount)}</strong>
              </article>
            </section>

            <section className="panel">
              <div className="panel__header panel__header--row">
                <div>
                  <h2>Последние расходы</h2>
                  <p>Последние записи за выбранный период.</p>
                </div>

                <button
                  className="small-button"
                  type="button"
                  onClick={() => setActiveSection('expenses')}
                >
                  Все расходы
                </button>
              </div>

              {latestExpenses.length === 0 ? (
                <div className="empty-state">За выбранный период расходов нет.</div>
              ) : (
                <div className="expense-list">
                  {latestExpenses.map((expense) => (
                    <article className="expense-card" key={expense.id}>
                      <div>
                        <p className="expense-card__date">{formatDate(expense.date)}</p>
                        <h3>{expense.description}</h3>
                        <p className="expense-card__category">{expense.category}</p>
                      </div>

                      <div className="expense-card__side">
                        <strong>{formatCurrency(expense.amount)}</strong>
                      </div>
                    </article>
                  ))}
                </div>
              )}
            </section>
          </>
        )}

        {activeSection === 'expenses' && (
          <>
            <section className="section-heading">
              <h2>Расходы</h2>
              <p>Добавление, редактирование и удаление расходов.</p>
            </section>

            <section className="panel">
              <div className="panel__header">
                <h2>Добавить расход</h2>
                <p>Заполните данные о расходе и сохраните запись.</p>
              </div>

              <form className="expense-form" onSubmit={handleSubmit}>
                <label>
                  <span>Дата</span>
                  <input
                    type="date"
                    value={date}
                    onChange={(event) => {
                      setDate(event.target.value)
                      event.currentTarget.blur()
                    }}
                  />
                </label>

                <label>
                  <span>Категория</span>
                  <select
                    value={category}
                    onChange={(event) => setCategory(event.target.value)}
                  >
                    {allCategories.map((currentCategory) => (
                      <option key={currentCategory} value={currentCategory}>
                        {currentCategory}
                      </option>
                    ))}
                  </select>
                </label>

                <label>
                  <span>Описание</span>
                  <input
                    type="text"
                    placeholder="Например: продукты в магазине"
                    value={description}
                    onChange={(event) => setDescription(event.target.value)}
                  />
                </label>

                <label>
                  <span>Сумма</span>
                  <input
                    type="text"
                    inputMode="decimal"
                    placeholder="0"
                    value={amount}
                    onChange={(event) =>
                      setAmount(sanitizeAmountInput(event.target.value))
                    }
                  />
                </label>

                <button className="button" type="submit">
                  Добавить расход
                </button>
              </form>
            </section>

            <section className="panel">
              <div className="panel__header">
                <h2>Список расходов</h2>
                <p>Список зависит от фильтра на главном экране.</p>
              </div>

              {sortedExpenses.length === 0 ? (
                <div className="empty-state">За выбранный период расходов нет.</div>
              ) : (
                <div className="expense-list">
                  {sortedExpenses.map((expense) => {
                    const isEditing = editingExpenseId === expense.id

                    return (
                      <article
                        className={`expense-card ${
                          isEditing ? 'expense-card--editing' : ''
                        }`}
                        key={expense.id}
                      >
                        {isEditing ? (
                          <form
                            className="expense-edit-form"
                            onSubmit={(event) => handleEditSubmit(event, expense.id)}
                          >
                            <label>
                              <span>Дата</span>
                              <input
                                type="date"
                                value={editDate}
                                onChange={(event) => {
                                  setEditDate(event.target.value)
                                  event.currentTarget.blur()
                                }}
                              />
                            </label>

                            <label>
                              <span>Категория</span>
                              <select
                                value={editCategory}
                                onChange={(event) =>
                                  setEditCategory(event.target.value)
                                }
                              >
                                {allCategories.map((currentCategory) => (
                                  <option key={currentCategory} value={currentCategory}>
                                    {currentCategory}
                                  </option>
                                ))}
                              </select>
                            </label>

                            <label>
                              <span>Описание</span>
                              <input
                                type="text"
                                value={editDescription}
                                onChange={(event) =>
                                  setEditDescription(event.target.value)
                                }
                              />
                            </label>

                            <label>
                              <span>Сумма</span>
                              <input
                                type="text"
                                inputMode="decimal"
                                value={editAmount}
                                onChange={(event) =>
                                  setEditAmount(sanitizeAmountInput(event.target.value))
                                }
                              />
                            </label>

                            <div className="expense-edit-form__actions">
                              <button className="button" type="submit">
                                Сохранить
                              </button>

                              <button
                                className="secondary-button"
                                type="button"
                                onClick={cancelEditing}
                              >
                                Отмена
                              </button>
                            </div>
                          </form>
                        ) : (
                          <>
                            <div>
                              <p className="expense-card__date">
                                {formatDate(expense.date)}
                              </p>
                              <h3>{expense.description}</h3>
                              <p className="expense-card__category">
                                {expense.category}
                              </p>
                            </div>

                            <div className="expense-card__side">
                              <strong>{formatCurrency(expense.amount)}</strong>

                              <div className="expense-card__actions">
                                <button
                                  className="edit-button"
                                  type="button"
                                  onClick={() => startEditing(expense)}
                                >
                                  Редактировать
                                </button>

                                <button
                                  className="delete-button"
                                  type="button"
                                  onClick={() => handleDelete(expense.id)}
                                >
                                  Удалить
                                </button>
                              </div>
                            </div>
                          </>
                        )}
                      </article>
                    )
                  })}
                </div>
              )}
            </section>
          </>
        )}

        {activeSection === 'reports' && (
          <section className="panel">
            <div className="panel__header">
              <h2>Отчет за период</h2>
              <p>
                Выберите даты и категорию, чтобы получить сводку и скачать Excel-отчет.
              </p>
            </div>

            <div className="report-filters">
              <label>
                <span>Дата начала</span>
                <input
                  type="date"
                  value={reportStartDate}
                  onChange={(event) => {
                    setReportStartDate(event.target.value)
                    event.currentTarget.blur()
                  }}
                />
              </label>

              <label>
                <span>Дата окончания</span>
                <input
                  type="date"
                  value={reportEndDate}
                  onChange={(event) => {
                    setReportEndDate(event.target.value)
                    event.currentTarget.blur()
                  }}
                />
              </label>

              <label>
                <span>Категория</span>
                <select
                  value={reportCategory}
                  onChange={(event) => setReportCategory(event.target.value)}
                >
                  <option value={ALL_CATEGORIES_OPTION}>
                    {ALL_CATEGORIES_OPTION}
                  </option>
                  {allCategories.map((currentCategory) => (
                    <option key={currentCategory} value={currentCategory}>
                      {currentCategory}
                    </option>
                  ))}
                </select>
              </label>
            </div>

            {isReportDateRangeInvalid ? (
              <div className="empty-state">
                Дата начала отчета не может быть позже даты окончания.
              </div>
            ) : (
              <>
                <section className="report-summary">
                  <article className="summary__card">
                    <span>Сумма отчета</span>
                    <strong>{formatCurrency(reportTotalAmount)}</strong>
                  </article>

                  <article className="summary__card">
                    <span>Операций</span>
                    <strong>{reportExpenses.length}</strong>
                  </article>

                  <article className="summary__card">
                    <span>Средний расход в день</span>
                    <strong>{formatCurrency(reportAveragePerDay)}</strong>
                  </article>
                </section>

                <button className="button" type="button" onClick={handleDownloadExcel}>
                  Скачать Excel-отчет
                </button>

                <div className="report-block">
                  <h3>Расходы по категориям</h3>

                  {reportCategorySummary.length === 0 ? (
                    <div className="empty-state">Нет данных за выбранный период.</div>
                  ) : (
                    <div className="table-wrapper">
                      <table>
                        <thead>
                          <tr>
                            <th>Категория</th>
                            <th>Операций</th>
                            <th>Сумма</th>
                            <th>Доля</th>
                          </tr>
                        </thead>
                        <tbody>
                          {reportCategorySummary.map((item) => (
                            <tr key={item.category}>
                              <td>{item.category}</td>
                              <td>{item.count}</td>
                              <td>{formatCurrency(item.amount)}</td>
                              <td>{item.percent.toFixed(1)}%</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                <div className="report-block">
                  <h3>Расходы по датам</h3>

                  {reportDailySummary.length === 0 ? (
                    <div className="empty-state">Нет данных за выбранный период.</div>
                  ) : (
                    <div className="table-wrapper">
                      <table>
                        <thead>
                          <tr>
                            <th>Дата</th>
                            <th>Операций</th>
                            <th>Сумма</th>
                          </tr>
                        </thead>
                        <tbody>
                          {reportDailySummary.map((item) => (
                            <tr key={item.date}>
                              <td>{formatDate(item.date)}</td>
                              <td>{item.count}</td>
                              <td>{formatCurrency(item.amount)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>

                <div className="report-block">
                  <h3>Детализация</h3>

                  {reportExpenses.length === 0 ? (
                    <div className="empty-state">Нет расходов для детализации.</div>
                  ) : (
                    <div className="table-wrapper">
                      <table>
                        <thead>
                          <tr>
                            <th>Дата</th>
                            <th>Категория</th>
                            <th>Описание</th>
                            <th>Сумма</th>
                          </tr>
                        </thead>
                        <tbody>
                          {reportExpenses.map((expense) => (
                            <tr key={expense.id}>
                              <td>{formatDate(expense.date)}</td>
                              <td>{expense.category}</td>
                              <td>{expense.description}</td>
                              <td>{formatCurrency(expense.amount)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              </>
            )}
          </section>
        )}

        {activeSection === 'categories' && (
          <section className="panel">
            <div className="panel__header">
              <h2>Категории</h2>
              <p>
                Добавляйте собственные категории. Они будут доступны в расходах,
                фильтрах, отчетах и резервных копиях.
              </p>
            </div>

            <form className="category-form" onSubmit={handleAddCategory}>
              <label>
                <span>Новая категория</span>
                <input
                  type="text"
                  placeholder="Например: Автомобиль"
                  value={newCategoryName}
                  onChange={(event) => setNewCategoryName(event.target.value)}
                />
              </label>

              <button className="button" type="submit">
                Добавить категорию
              </button>
            </form>

            <div className="category-list">
              {allCategories.map((currentCategory) => {
                const isSystemCategory = hasCategoryName(
                  DEFAULT_CATEGORIES,
                  currentCategory,
                )
                const isEditingCategory = editingCategoryName === currentCategory

                return (
                  <article className="category-card" key={currentCategory}>
                    {isEditingCategory ? (
                      <form
                        className="category-edit-form"
                        onSubmit={handleRenameCategory}
                      >
                        <label>
                          <span>Название категории</span>
                          <input
                            type="text"
                            value={editedCategoryName}
                            onChange={(event) =>
                              setEditedCategoryName(event.target.value)
                            }
                          />
                        </label>

                        <div className="category-card__actions">
                          <button className="edit-button" type="submit">
                            Сохранить
                          </button>

                          <button
                            className="delete-button"
                            type="button"
                            onClick={cancelCategoryEditing}
                          >
                            Отмена
                          </button>
                        </div>
                      </form>
                    ) : (
                      <>
                        <div>
                          <h3>{currentCategory}</h3>
                          <p>
                            {isSystemCategory
                              ? 'Системная категория'
                              : 'Пользовательская категория'}
                          </p>
                        </div>

                        {!isSystemCategory && (
                          <div className="category-card__actions">
                            <button
                              className="edit-button"
                              type="button"
                              onClick={() => startCategoryEditing(currentCategory)}
                            >
                              Переименовать
                            </button>

                            <button
                              className="delete-button"
                              type="button"
                              onClick={() => handleDeleteCategory(currentCategory)}
                            >
                              Удалить
                            </button>
                          </div>
                        )}
                      </>
                    )}
                  </article>
                )
              })}
            </div>
          </section>
        )}

        {activeSection === 'settings' && (
          <section className="panel">
            <div className="panel__header">
              <h2>Настройки</h2>
              <p>
                Здесь можно создать резервную копию данных или восстановить их из файла.
              </p>
            </div>

            <div className="settings-note">
              <h3>Где хранятся данные?</h3>
              <p>
                Расходы и категории хранятся локально на этом устройстве. У каждого
                телефона или компьютера будут свои данные.
              </p>
            </div>

            <div className="backup-actions">
              <button className="button" type="button" onClick={handleCreateBackup}>
                Скачать резервную копию
              </button>

              <label className="restore-button">
                Восстановить из файла
                <input
                  type="file"
                  accept="application/json,.json"
                  onChange={handleRestoreBackup}
                />
              </label>
            </div>
          </section>
        )}
      </section>
    </main>
  )
}

export default App