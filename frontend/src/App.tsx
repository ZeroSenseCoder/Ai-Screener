import { BrowserRouter, Routes, Route } from 'react-router-dom'
import Navbar from './components/common/Navbar'
import MacroDashboard from './components/macro/MacroDashboard'
import ScreenerPage from './pages/ScreenerPage'
import StockDetailPage from './pages/StockDetailPage'
import ProScreenerPage from './pages/ProScreenerPage'
import ValuationPage from './pages/ValuationPage'

export default function App() {
  return (
    <BrowserRouter>
      <div className="flex flex-col h-screen overflow-hidden">
        <Navbar />
        <MacroDashboard />
        <main className="flex-1 min-h-0 overflow-hidden">
          <Routes>
            <Route path="/" element={<ScreenerPage />} />
            <Route path="/pro-screener" element={<ProScreenerPage />} />
            <Route path="/stock/:symbol" element={<StockDetailPage />} />
            <Route path="/stock/:symbol/valuation" element={<ValuationPage />} />
          </Routes>
        </main>
      </div>
    </BrowserRouter>
  )
}
