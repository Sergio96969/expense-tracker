import './App.css'

function App() {
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

        <section className="summary">
          <article className="summary__card">
            <span>Расходы за месяц</span>
            <strong>0 ₽</strong>
          </article>

          <article className="summary__card">
            <span>Количество записей</span>
            <strong>0</strong>
          </article>

          <article className="summary__card">
            <span>Средний расход</span>
            <strong>0 ₽</strong>
          </article>
        </section>

        <section className="panel">
          <div className="panel__header">
            <h2>Добавить расход</h2>
            <p>Позже здесь появится форма добавления расходов.</p>
          </div>

          <button className="button" type="button">
            Добавить первый расход
          </button>
        </section>
      </section>
    </main>
  )
}

export default App