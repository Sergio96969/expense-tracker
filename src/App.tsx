import { useEffect, useMemo, useState } from 'react'
import type { FormEvent } from 'react'
import './App.css'

type Expense = {
  id: string
  date: string
  category: string
  description: string
  amount: number
}

const STORAGE_KEY = 'expense-tracker-expenses'

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
  'Другое',
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

function getTodayDate() {
  return new Date().toISOString().slice(0, 10)
}

function getCurrentMonth() {
  return new Date().toISOString().slice(5, 7)
}

function getCurrentYear() {
  return String(new Date().getFullYear())
}

function formatCurrency(value: number) {
  return new Intl.NumberFormat('ru-RU', {
    style: 'currency',
    currency: 'RUB',
    maximumFractionDigits: 0,
  }).format(value)
}

function App() {
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

  const [date, setDate] = useState(getTodayDate())
  const [category, setCategory] = useState(DEFAULT_CATEGORIES[0])
  const [description, setDescription] = useState('')
  const [amount, setAmount] = useState('')
  const [message, setMessage] = useState('')
  const [selectedMonth, setSelectedMonth] = useState(getCurrentMonth())
  const [selectedYear, setSelectedYear] = useState(getCurrentYear())
  const [selectedCategory, setSelectedCategory] = useState('Все категории')

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(expenses))
  }, [expenses])

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
        selectedCategory === 'Все категории' || expense.category === selectedCategory

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

  function handleDelete(id: string) {
    const shouldDelete = window.confirm('Удалить этот расход?')

    if (!shouldDelete) {
      return
    }

    setExpenses((currentExpenses) =>
      currentExpenses.filter((expense) => expense.id !== id),
    )

    setMessage('Расход удален')
  }

  return (
    <main className="app">
      <section className="app__container">
        <header className="app__header">
          <div>
            <p className="app__eyebrow">Личный финансовый помощник</p>
            <h1>Учет расходов</h1>
            <p className="app__description">
              Простое приложение для фиксации расходов, анализа трат и формирования отчетов.
            </p>
          </div>
        </header>

        {message && (
          <div className="toast">
            <span>{message}</span>
            <button type="button" onClick={() => setMessage('')}>
              ×
            </button>
          </div>
        )}

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
              <option value="Все категории">Все категории</option>
              {DEFAULT_CATEGORIES.map((currentCategory) => (
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
                {DEFAULT_CATEGORIES.map((currentCategory) => (
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
                onChange={(event) => {
                  const onlyNumbers = event.target.value.replace(/[^\d.,]/g, '')
                  const normalizedValue = onlyNumbers.replace(',', '.')
                  setAmount(normalizedValue)
                }}
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
            <p>Все добавленные расходы сохраняются локально в браузере.</p>
          </div>

          {sortedExpenses.length === 0 ? (
            <div className="empty-state">
              За выбранный период расходов нет.
            </div>
          ) : (
            <div className="expense-list">
              {sortedExpenses.map((expense) => (
                <article className="expense-card" key={expense.id}>
                  <div>
                    <p className="expense-card__date">{expense.date}</p>
                    <h3>{expense.description}</h3>
                    <p className="expense-card__category">{expense.category}</p>
                  </div>

                  <div className="expense-card__side">
                    <strong>{formatCurrency(expense.amount)}</strong>
                    <button
                      className="delete-button"
                      type="button"
                      onClick={() => handleDelete(expense.id)}
                    >
                      Удалить
                    </button>
                  </div>
                </article>
              ))}
            </div>
          )}
        </section>
      </section>
    </main>
  )
}

export default App