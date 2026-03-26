import axios from 'axios'

const api = axios.create({ baseURL: '/api/v1' })

export interface Stock {
  symbol: string
  company_name: string
  exchange: string
  sector: string
  industry: string
  yf_symbol: string
  last_price?: number
  daily_return?: number
  return_1m?: number
  return_3m?: number
  return_1y?: number
  market_cap?: number
  pe_ratio?: number
  beta?: number
  rsi_14?: number
  sma_20?: number
  sma_50?: number
  sma_200?: number
  macd?: number
  macd_signal?: number
  avg_volume_20d?: number
  max_drawdown_52w?: number
}

export interface ScreenerRequest {
  exchanges?: string[]
  sectors?: string[]
  min_price?: number
  max_price?: number
  min_market_cap?: number
  max_market_cap?: number
  min_pe?: number
  max_pe?: number
  min_beta?: number
  max_beta?: number
  min_rsi?: number
  max_rsi?: number
  price_above_sma20?: boolean
  price_above_sma50?: boolean
  price_above_sma200?: boolean
  sma50_above_sma200?: boolean
  macd_bullish?: boolean
  macd_bearish?: boolean
  min_daily_return?: number
  max_daily_return?: number
  min_return_1m?: number
  max_return_1m?: number
  min_return_3m?: number
  max_return_3m?: number
  min_return_1y?: number
  max_return_1y?: number
  min_avg_volume?: number
  max_drawdown_threshold?: number
  sort_by?: string
  sort_asc?: boolean
  page?: number
  page_size?: number
}

export interface ScreenerResponse {
  results: Stock[]
  total: number
  page: number
  page_size: number
  pages: number
}

export const screenerApi = {
  screen: (req: ScreenerRequest) =>
    api.post<ScreenerResponse>('/screener/screen', req).then(r => r.data),
  meta: () => api.get<{ sectors: string[]; industries: string[]; exchanges: string[] }>('/screener/meta').then(r => r.data),
  sectorSummary: () => api.get('/screener/sector-summary').then(r => r.data),
}

export const universeApi = {
  stats: () => api.get('/universe/stats').then(r => r.data),
  sectors: () => api.get('/universe/sectors/list').then(r => r.data),
  sectorStocks: (name: string) => api.get(`/universe/sectors/${encodeURIComponent(name)}`).then(r => r.data),
  allStocks: (params?: { exchange?: string; sector?: string; search?: string; page?: number }) =>
    api.get('/universe/stocks', { params }).then(r => r.data),
}

export const stockApi = {
  search: (q: string) => api.get<Stock[]>('/stocks/search', { params: { q } }).then(r => r.data),
  ohlcv: (symbol: string, timeframe = '1D') =>
    api.get(`/stocks/${symbol}/ohlcv`, { params: { timeframe } }).then(r => r.data),
  indicators: (symbol: string) => api.get(`/stocks/${symbol}/indicators`).then(r => r.data),
  fundamentals: (symbol: string) => api.get(`/stocks/${symbol}/fundamentals`).then(r => r.data),
  news: (symbol: string, days = 7) => api.get(`/stocks/${symbol}/news`, { params: { days } }).then(r => r.data),
  summary: (symbol: string) => api.get(`/stocks/${symbol}/summary`).then(r => r.data),
  oi: (symbol: string) => api.get(`/stocks/${symbol}/oi`).then(r => r.data),
}

// ── Pro Screener ───────────────────────────────────────────────────────────────

export interface FilterCondition {
  filter_id: string
  min_val?: number | null
  max_val?: number | null
  weight: number       // decimal 0–1; all active conditions must sum to 1
  required: boolean    // hard filter (exclude if fails) vs soft (score only)
}

export interface ProScreenRequest {
  sector?: string | null
  conditions: FilterCondition[]
  score_mode: boolean
  sort_by?: string
  sort_asc?: boolean
  page?: number
  page_size?: number
}

export interface ProStockResult {
  symbol: string
  company_name: string
  exchange: string
  sector: string
  score: number
  last_price?: number
  market_cap?: number
  pe_ratio?: number
  beta?: number
  rsi_14?: number
  macd_hist?: number
  max_drawdown_52w?: number
  daily_return?: number
  return_5d?: number
  return_1m?: number
  return_3m?: number
  return_1y?: number
  avg_volume_20d?: number
}

export interface ProScreenResponse {
  total: number
  page: number
  page_size: number
  sector: string | null
  score_mode: boolean
  conditions: number
  results: ProStockResult[]
}

export interface CatalogFilter {
  id: string
  label: string
  type: 'range' | 'bool'
  unit: string
  available: boolean
}

export interface CatalogCategory {
  category: string
  label: string
  filters: CatalogFilter[]
}

export const proScreenerApi = {
  catalog: () => api.get<CatalogCategory[]>('/pro-screener/catalog').then(r => r.data),
  screen: (req: ProScreenRequest) =>
    api.post<ProScreenResponse>('/pro-screener/screen', req).then(r => r.data),
  sectors: () => api.get<{ sector: string; count: number }[]>('/pro-screener/sectors').then(r => r.data),
}

export const macroApi = {
  overview: () => api.get('/macro/overview').then(r => r.data),
  indices: () => api.get('/macro/indices').then(r => r.data),
  forex: () => api.get('/macro/forex').then(r => r.data),
  commodities: () => api.get('/macro/commodities').then(r => r.data),
  fiiDii: () => api.get('/macro/fii-dii').then(r => r.data),
}

export const valuationApi = {
  inputs: (symbol: string) => api.get(`/valuation/${symbol}/inputs`).then(r => r.data),
  run: (symbol: string, model: string, params: Record<string, any>) =>
    api.post(`/valuation/${symbol}/run`, { model, params }).then(r => r.data),
}

export default api
