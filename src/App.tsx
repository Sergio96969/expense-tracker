import { FormEvent, useEffect, useMemo, useState } from 'react'
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

function getTodayDate() {
  return new Date().toISOString().slice(0, 10)
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

  const totalAmount = useMemo(() => {
    return expenses.reduce((sum, expense) => sum + expense.amount, 0)
  }, [expenses])

  const averageAmount = expenses.length > 0 ? totalAmount / expenses.length : 0

  const sortedExpenses = useMemo(() => {
    return [...expenses].sort((a, b) => b.date.localeCompare(a.date))
  }, [expenses])

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

        <section className="summary">
          <article className="summary__card">
            <span>Расходы всего</span>
            <strong>{formatCurrency(totalAmount)}</strong>
          </article>

          <article className="summary__card">
            <span>Количество записей</span>
            <strong>{expenses.length}</strong>
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
/>            </label>

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
              Пока расходов нет. Добавьте первую запись.
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