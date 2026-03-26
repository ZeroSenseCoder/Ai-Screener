import { useEffect, useState, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import * as XLSX from 'xlsx'
import api from '../services/api'

// ── Types ─────────────────────────────────────────────────────────────────────
interface BsField {
  key: string
  label: string
}

interface ParamDef {
  key: string
  label: string
  type?: 'number' | 'select' | 'percent'
  required?: boolean
  default?: number
  min?: number
  max?: number
  step?: number
  options?: { value: string; label: string }[]
  hint?: string
}

interface ModelDef {
  id: string
  label: string
  category: string
  description: string
  bsFields: BsField[]
  params: ParamDef[]
  backendModel?: string
}

// ── Model catalog ─────────────────────────────────────────────────────────────
const MODELS: ModelDef[] = [
  // ── Intrinsic Valuation ────────────────────────────────────────────────────
  {
    id: 'dcf_fcff',
    label: 'DCF – Free Cash Flow to Firm',
    category: 'Intrinsic Valuation',
    description: 'Discounts projected FCFF at WACC to derive Enterprise Value, then subtracts net debt for equity value.',
    bsFields: [
      { key: 'revenue', label: 'Revenue' },
      { key: 'ebitda', label: 'EBITDA' },
      { key: 'operating_cf', label: 'Operating Cash Flow' },
      { key: 'fcf', label: 'Free Cash Flow' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash & Equivalents' },
      { key: 'depreciation', label: 'Depreciation & Amortization' },
    ],
    params: [
      { key: 'fcf', label: 'Free Cash Flow (₹)', required: true },
      { key: 'growth_stage1', label: 'FCF Growth Rate', type: 'percent', required: true, default: 12, min: 0, max: 40, step: 0.5 },
      { key: 'terminal_growth', label: 'Terminal Growth Rate', type: 'percent', required: true, default: 5, min: 0, max: 10, step: 0.5 },
      { key: 'wacc', label: 'WACC', type: 'percent', required: true, default: 12, min: 5, max: 25, step: 0.5 },
      { key: 'years', label: 'Projection Period (yrs)', default: 5, min: 3, max: 15, step: 1 },
      { key: 'total_debt', label: 'Total Debt (₹)' },
      { key: 'cash', label: 'Cash (₹)' },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },
  {
    id: 'dcf_fcfe',
    label: 'DCF – Free Cash Flow to Equity',
    category: 'Intrinsic Valuation',
    description: 'Discounts FCFE directly at cost of equity. Bypasses WACC — suitable when debt is stable.',
    bsFields: [
      { key: 'fcfe', label: 'FCFE' },
      { key: 'operating_cf', label: 'Operating CF' },
      { key: 'capex', label: 'Capex' },
      { key: 'shares', label: 'Shares Outstanding' },
    ],
    params: [
      { key: 'fcfe', label: 'FCFE (₹)', required: true },
      { key: 'growth_stage1', label: 'FCFE Growth Rate', type: 'percent', required: true, default: 12, min: 0, max: 40, step: 0.5 },
      { key: 'terminal_growth', label: 'Terminal Growth', type: 'percent', required: true, default: 5, min: 0, max: 10, step: 0.5 },
      { key: 'cost_of_equity', label: 'Cost of Equity', type: 'percent', required: true, default: 14, min: 6, max: 25, step: 0.5 },
      { key: 'years', label: 'Projection Years', default: 5, min: 3, max: 15, step: 1 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },
  {
    id: 'dcf_multistage',
    label: 'Multi-Stage DCF (3 Stages)',
    category: 'Intrinsic Valuation',
    description: 'Three-stage DCF: high growth → transition → terminal. Best for companies with changing growth profiles.',
    bsFields: [
      { key: 'fcf', label: 'Current FCF' },
      { key: 'revenue', label: 'Revenue' },
      { key: 'ebitda', label: 'EBITDA' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'fcf', label: 'Current FCF (₹)', required: true },
      { key: 'growth_stage1', label: 'Stage 1 Growth', type: 'percent', required: true, default: 20, min: 0, max: 50, step: 1 },
      { key: 'growth_stage2', label: 'Stage 2 Growth', type: 'percent', required: true, default: 12, min: 0, max: 30, step: 1 },
      { key: 'terminal_growth', label: 'Terminal Growth', type: 'percent', required: true, default: 5, min: 0, max: 10, step: 0.5 },
      { key: 'wacc', label: 'WACC', type: 'percent', required: true, default: 12, min: 5, max: 25, step: 0.5 },
      { key: 'stage1_years', label: 'Stage 1 Years', default: 5, min: 1, max: 10, step: 1 },
      { key: 'stage2_years', label: 'Stage 2 Years', default: 5, min: 1, max: 10, step: 1 },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },

  // ── Dividend Discount ──────────────────────────────────────────────────────
  {
    id: 'gordon_growth',
    label: 'Gordon Growth Model (DDM)',
    category: 'Dividend Discount',
    description: 'Classic one-stage DDM. Values dividend-paying stocks via perpetuity formula: P = D1 / (ke − g).',
    bsFields: [
      { key: 'dps', label: 'Dividend Per Share (₹)' },
      { key: 'eps', label: 'Earnings Per Share' },
    ],
    params: [
      { key: 'dps', label: 'Dividend Per Share (₹)', required: true },
      { key: 'cost_of_equity', label: 'Cost of Equity', type: 'percent', required: true, default: 14, min: 6, max: 25, step: 0.5 },
      { key: 'terminal_growth', label: 'Dividend Growth Rate', type: 'percent', required: true, default: 5, min: 0, max: 12, step: 0.5 },
    ],
  },
  {
    id: 'ddm_multistage',
    label: 'Multi-Stage DDM',
    category: 'Dividend Discount',
    description: 'Two-stage dividend model: explicit high-growth period followed by perpetuity at terminal growth rate.',
    bsFields: [
      { key: 'dps', label: 'Current DPS' },
      { key: 'eps', label: 'EPS' },
      { key: 'roe', label: 'Return on Equity' },
    ],
    params: [
      { key: 'dps', label: 'Current DPS (₹)', required: true },
      { key: 'growth_stage1', label: 'Stage 1 Dividend Growth', type: 'percent', required: true, default: 10, min: 0, max: 30, step: 1 },
      { key: 'terminal_growth', label: 'Terminal Growth', type: 'percent', required: true, default: 5, min: 0, max: 10, step: 0.5 },
      { key: 'cost_of_equity', label: 'Cost of Equity', type: 'percent', required: true, default: 14, min: 6, max: 25, step: 0.5 },
      { key: 'years', label: 'Stage 1 Years', default: 5, min: 2, max: 10, step: 1 },
    ],
  },

  // ── Residual Income ────────────────────────────────────────────────────────
  {
    id: 'residual_income',
    label: 'Residual Income Model (RIM)',
    category: 'Residual Income',
    description: 'Value = Book Value + PV of excess returns above equity cost. Useful when dividends are low or zero.',
    bsFields: [
      { key: 'book_value_per_share', label: 'Book Value Per Share' },
      { key: 'eps', label: 'EPS' },
      { key: 'roe', label: 'Return on Equity' },
      { key: 'net_income', label: 'Net Income' },
    ],
    params: [
      { key: 'book_value_per_share', label: 'Book Value Per Share (₹)', required: true },
      { key: 'roe', label: 'ROE', type: 'percent', required: true, default: 15, min: 1, max: 40, step: 1 },
      { key: 'cost_of_equity', label: 'Cost of Equity', type: 'percent', required: true, default: 14, min: 6, max: 25, step: 0.5 },
      { key: 'terminal_growth', label: 'Terminal Growth', type: 'percent', default: 5, min: 0, max: 10, step: 0.5 },
      { key: 'years', label: 'Forecast Years', default: 5, min: 3, max: 15, step: 1 },
    ],
  },

  // ── Relative Valuation ─────────────────────────────────────────────────────
  {
    id: 'trading_comps',
    label: 'Trading Comparables',
    category: 'Relative Valuation',
    description: 'Values the stock using median P/E, P/B, P/S multiples of sector peers sourced from the universe.',
    bsFields: [
      { key: 'market_cap', label: 'Market Cap' },
      { key: 'enterprise_value', label: 'Enterprise Value' },
      { key: 'revenue', label: 'Revenue' },
      { key: 'ebitda', label: 'EBITDA' },
      { key: 'net_income', label: 'Net Income' },
      { key: 'book_value_total', label: 'Book Value' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
      { key: 'eps', label: 'EPS' },
    ],
    params: [],
  },
  {
    id: 'precedent_transactions',
    label: 'Precedent Transactions',
    category: 'Relative Valuation',
    description: 'M&A deal implied value using transaction EV/EBITDA multiples plus control premium and synergies.',
    bsFields: [
      { key: 'revenue', label: 'Revenue' },
      { key: 'ebitda', label: 'EBITDA' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'ev_ebitda_multiple', label: 'Transaction EV/EBITDA', required: true, default: 10, min: 3, max: 25, step: 0.5 },
      { key: 'deal_premium', label: 'Deal Premium', type: 'percent', required: true, default: 20, min: 0, max: 60, step: 5 },
      { key: 'synergies_pct', label: 'Synergies (% of EBITDA)', type: 'percent', default: 5, min: 0, max: 30, step: 2.5 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
  },
  {
    id: 'peg',
    label: 'PEG Ratio Model',
    category: 'Relative Valuation',
    description: 'Fair value using PEG = 1: Price = EPS × Expected Growth %. Simple growth-adjusted P/E model.',
    bsFields: [
      { key: 'eps', label: 'EPS' },
      { key: 'earnings_growth', label: 'Earnings Growth Rate' },
    ],
    params: [
      { key: 'eps', label: 'EPS (₹)', required: true },
      { key: 'earnings_growth_pct', label: 'EPS Growth (%)', required: true, default: 15, min: 1, max: 50, step: 1 },
      { key: 'target_peg', label: 'Target PEG Ratio', default: 1.0, min: 0.5, max: 2.0, step: 0.1 },
    ],
  },
  {
    id: 'revenue_multiple',
    label: 'Revenue Multiple',
    category: 'Relative Valuation',
    description: 'EV = Revenue × EV/Revenue multiple. Useful for high-growth or pre-profit companies.',
    bsFields: [
      { key: 'revenue', label: 'Revenue' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'ev_revenue_multiple', label: 'EV/Revenue Multiple', required: true, default: 3, min: 0.5, max: 20, step: 0.5 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
  },

  // ── Asset-Based ────────────────────────────────────────────────────────────
  {
    id: 'nav',
    label: 'Net Asset Value (NAV)',
    category: 'Asset-Based',
    description: 'Equity value = Adjusted Assets − Liabilities. Useful for asset-heavy companies, holding cos, REITs.',
    bsFields: [
      { key: 'total_assets', label: 'Total Assets' },
      { key: 'total_liabilities', label: 'Total Liabilities' },
      { key: 'fixed_assets', label: 'Fixed Assets' },
    ],
    params: [
      { key: 'total_assets', label: 'Total Assets (₹)', required: true },
      { key: 'total_liabilities', label: 'Total Liabilities (₹)', required: true },
      { key: 'goodwill', label: 'Goodwill to Exclude (₹)', default: 0 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },
  {
    id: 'liquidation',
    label: 'Liquidation Value',
    category: 'Asset-Based',
    description: 'Floor valuation — recoverable asset value after applying haircuts to each asset class.',
    bsFields: [
      { key: 'total_assets', label: 'Total Assets' },
      { key: 'total_liabilities', label: 'Total Liabilities' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'cash_rate', label: 'Cash Recovery %', type: 'percent', default: 100, min: 80, max: 100, step: 5 },
      { key: 'receivables_rate', label: 'Receivables Recovery %', type: 'percent', default: 85, min: 40, max: 100, step: 5 },
      { key: 'inventory_rate', label: 'Inventory Recovery %', type: 'percent', default: 50, min: 20, max: 90, step: 5 },
      { key: 'ppe_rate', label: 'PP&E Recovery %', type: 'percent', default: 60, min: 20, max: 90, step: 5 },
      { key: 'other_rate', label: 'Other Assets Recovery %', type: 'percent', default: 25, min: 5, max: 60, step: 5 },
    ],
  },
  {
    id: 'replacement_cost',
    label: 'Replacement Cost',
    category: 'Asset-Based',
    description: 'Values the firm based on cost to recreate its asset base today. Reflects economic replacement value.',
    bsFields: [
      { key: 'total_assets', label: 'Total Assets' },
      { key: 'fixed_assets', label: 'Fixed / PP&E Assets' },
      { key: 'total_liabilities', label: 'Total Liabilities' },
    ],
    params: [
      { key: 'rebuild_multiplier', label: 'Rebuild Cost Multiplier', required: true, default: 1.2, min: 0.8, max: 3.0, step: 0.1, hint: 'Cost to rebuild assets today vs book' },
      { key: 'depreciation_adj', label: 'Depreciation Adjustment', type: 'percent', default: 30, min: 0, max: 70, step: 5, hint: '% of replacement cost already depreciated' },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },

  // ── Earnings-Based ─────────────────────────────────────────────────────────
  {
    id: 'capitalized_earnings',
    label: 'Capitalized Earnings',
    category: 'Earnings-Based',
    description: 'Capitalizes normalized EPS at required return rate. Simple and intuitive for stable earners.',
    bsFields: [
      { key: 'net_income', label: 'Net Income' },
      { key: 'eps', label: 'EPS' },
      { key: 'revenue', label: 'Revenue' },
    ],
    params: [
      { key: 'eps', label: 'EPS (₹)', required: true },
      { key: 'required_return', label: 'Required Return', type: 'percent', required: true, default: 14, min: 6, max: 25, step: 0.5 },
      { key: 'eps_growth', label: 'EPS Normalization Growth', type: 'percent', default: 5, min: 0, max: 20, step: 1 },
    ],
  },
  {
    id: 'excess_earnings',
    label: 'Excess Earnings Model',
    category: 'Earnings-Based',
    description: 'Value = Tangible assets + PV of earnings in excess of a fair return on assets.',
    bsFields: [
      { key: 'net_income', label: 'Net Income' },
      { key: 'total_assets', label: 'Total Assets' },
      { key: 'book_value_total', label: 'Book Value of Equity' },
    ],
    params: [
      { key: 'fair_return_rate', label: 'Fair Return on Assets', type: 'percent', required: true, default: 8, min: 3, max: 20, step: 0.5, hint: 'Expected normal return on total assets' },
      { key: 'discount_rate', label: 'Discount Rate', type: 'percent', required: true, default: 12, min: 5, max: 25, step: 0.5 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },

  // ── Economic Profit ────────────────────────────────────────────────────────
  {
    id: 'eva',
    label: 'Economic Value Added (EVA)',
    category: 'Economic Profit',
    description: 'Firm value = Invested Capital + PV of all future EVAs (NOPAT − WACC × IC).',
    bsFields: [
      { key: 'nopat', label: 'NOPAT' },
      { key: 'invested_capital', label: 'Invested Capital' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'nopat', label: 'NOPAT (₹)', required: true },
      { key: 'invested_capital', label: 'Invested Capital (₹)', required: true },
      { key: 'wacc', label: 'WACC', type: 'percent', required: true, default: 12, min: 5, max: 25, step: 0.5 },
      { key: 'growth_stage1', label: 'NOPAT Growth Rate', type: 'percent', default: 8, min: 0, max: 25, step: 1 },
      { key: 'terminal_growth', label: 'Terminal Growth', type: 'percent', default: 4, min: 0, max: 10, step: 0.5 },
      { key: 'years', label: 'Forecast Years', default: 5, min: 3, max: 15, step: 1 },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },
  {
    id: 'cfroi',
    label: 'Cash Flow Return on Investment (CFROI)',
    category: 'Economic Profit',
    description: 'Compares CFROI (operating CF / asset base) vs required return. Spread drives value creation.',
    bsFields: [
      { key: 'operating_cf', label: 'Operating Cash Flow' },
      { key: 'total_assets', label: 'Total Asset Base' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'asset_life', label: 'Asset Life (years)', required: true, default: 10, min: 3, max: 30, step: 1 },
      { key: 'required_return', label: 'Required Return', type: 'percent', required: true, default: 12, min: 5, max: 25, step: 0.5 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },

  // ── Private Equity ─────────────────────────────────────────────────────────
  {
    id: 'lbo',
    label: 'Leveraged Buyout (LBO)',
    category: 'Private Equity',
    description: 'PE-style LBO model: entry/exit multiples, leverage, debt paydown, computes IRR and MOIC.',
    bsFields: [
      { key: 'ebitda', label: 'EBITDA' },
      { key: 'revenue', label: 'Revenue' },
      { key: 'total_debt', label: 'Existing Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'ebitda', label: 'EBITDA (₹)', required: true },
      { key: 'entry_multiple', label: 'Entry EV/EBITDA', required: true, default: 8, min: 3, max: 20, step: 0.5 },
      { key: 'exit_multiple', label: 'Exit EV/EBITDA', required: true, default: 10, min: 3, max: 25, step: 0.5 },
      { key: 'debt_ratio', label: 'Leverage (Debt/EV)', type: 'percent', required: true, default: 60, min: 20, max: 80, step: 5 },
      { key: 'interest_rate', label: 'Interest Rate', type: 'percent', default: 9, min: 5, max: 18, step: 0.5 },
      { key: 'ebitda_growth', label: 'EBITDA Growth', type: 'percent', default: 8, min: 0, max: 25, step: 1 },
      { key: 'hold_years', label: 'Hold Period (yrs)', default: 5, min: 3, max: 10, step: 1 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },

  // ── Option-Based ───────────────────────────────────────────────────────────
  {
    id: 'black_scholes',
    label: 'Black-Scholes Option Pricing',
    category: 'Option-Based',
    description: 'Standard B-S model for European call/put options. Uses live stock price as spot.',
    bsFields: [
      { key: 'price', label: 'Current Stock Price (Spot)' },
    ],
    params: [
      { key: 'strike', label: 'Strike Price (₹)', required: true },
      { key: 'volatility', label: 'Annualized Volatility', type: 'percent', required: true, default: 30, min: 5, max: 100, step: 1 },
      { key: 'risk_free', label: 'Risk-Free Rate', type: 'percent', default: 7.1, min: 1, max: 12, step: 0.1 },
      { key: 'time_years', label: 'Time to Expiry (yrs)', default: 1, min: 0.1, max: 5, step: 0.1 },
      { key: 'option_type', label: 'Option Type', type: 'select', options: [{ value: 'call', label: 'Call' }, { value: 'put', label: 'Put' }] },
    ],
  },
  {
    id: 'real_options',
    label: 'Real Options Valuation',
    category: 'Option-Based',
    description: 'Values managerial flexibility (invest/expand/abandon) using Black-Scholes framework.',
    backendModel: 'black_scholes',
    bsFields: [
      { key: 'operating_cf', label: 'Project Cash Flows (Operating CF)' },
      { key: 'capex', label: 'Investment Required (Capex)' },
    ],
    params: [
      { key: 'strike', label: 'Investment Cost (₹)', required: true, hint: 'Cost to exercise the option (capex)' },
      { key: 'volatility', label: 'Project Volatility', type: 'percent', required: true, default: 40, min: 10, max: 100, step: 5, hint: 'Volatility of underlying project value' },
      { key: 'risk_free', label: 'Risk-Free Rate', type: 'percent', default: 7.1, min: 1, max: 12, step: 0.1 },
      { key: 'time_years', label: 'Time Horizon (yrs)', default: 3, min: 0.5, max: 10, step: 0.5 },
      { key: 'option_type', label: 'Option Type', type: 'select', options: [{ value: 'call', label: 'Call (Option to invest)' }, { value: 'put', label: 'Put (Option to abandon)' }] },
    ],
  },

  // ── Sum of Parts ───────────────────────────────────────────────────────────
  {
    id: 'sum_of_parts',
    label: 'Sum of the Parts',
    category: 'Sum of Parts',
    description: 'Values each business segment separately and sums them. Enter segments in the params below.',
    bsFields: [
      { key: 'ebitda', label: 'Total EBITDA' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
      { key: 'revenue', label: 'Total Revenue' },
    ],
    params: [
      { key: 'total_debt', label: 'Total Debt (₹)' },
      { key: 'cash', label: 'Cash (₹)' },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },

  // ── Industry-Specific ──────────────────────────────────────────────────────
  {
    id: 'pb_banks',
    label: 'P/B Model (Banks & NBFCs)',
    category: 'Industry-Specific',
    description: 'Gordon-growth implied P/B = (ROE − g) / (ke − g). Standard bank valuation framework.',
    bsFields: [
      { key: 'book_value_per_share', label: 'Book Value Per Share' },
      { key: 'roe', label: 'Return on Equity' },
      { key: 'net_income', label: 'Net Income' },
    ],
    params: [
      { key: 'book_value_per_share', label: 'Book Value Per Share (₹)', required: true },
      { key: 'roe', label: 'ROE', type: 'percent', required: true, default: 15, min: 5, max: 30, step: 1 },
      { key: 'cost_of_equity', label: 'Cost of Equity', type: 'percent', required: true, default: 14, min: 6, max: 25, step: 0.5 },
      { key: 'terminal_growth', label: 'Dividend Growth Rate', type: 'percent', default: 7, min: 2, max: 12, step: 0.5 },
    ],
  },
  {
    id: 'cap_rate',
    label: 'Cap Rate (Real Estate / REITs)',
    category: 'Industry-Specific',
    description: 'Property value = NOI / Cap Rate. Standard real estate and REIT valuation metric.',
    bsFields: [
      { key: 'noi', label: 'Net Operating Income' },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
    params: [
      { key: 'noi', label: 'Net Operating Income (₹)', required: true },
      { key: 'cap_rate', label: 'Market Cap Rate', type: 'percent', required: true, default: 7, min: 4, max: 15, step: 0.25 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
      { key: 'total_debt', label: 'Total Debt' },
      { key: 'cash', label: 'Cash' },
    ],
  },
  {
    id: 'user_based',
    label: 'User-Based Valuation (SaaS / Platform)',
    category: 'Industry-Specific',
    description: 'Projects platform value from user count, revenue per user, growth and churn dynamics.',
    bsFields: [
      { key: 'revenue', label: 'Total Revenue' },
      { key: 'net_income', label: 'Net Income' },
    ],
    params: [
      { key: 'users', label: 'Total Users / Subscribers', required: true, hint: 'From annual report disclosure' },
      { key: 'revenue_per_user', label: 'Revenue Per User (₹/yr)', required: true, hint: 'Avg annual revenue per user' },
      { key: 'user_growth', label: 'User Growth Rate', type: 'percent', required: true, default: 15, min: 0, max: 100, step: 5 },
      { key: 'churn_rate', label: 'Annual Churn Rate', type: 'percent', required: true, default: 5, min: 0, max: 50, step: 1 },
      { key: 'discount_rate', label: 'Discount Rate', type: 'percent', default: 14, min: 6, max: 30, step: 0.5 },
      { key: 'years', label: 'Projection Years', default: 5, min: 3, max: 10, step: 1 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },
  {
    id: 'vc_method',
    label: 'VC / Venture Capital Method',
    category: 'Industry-Specific',
    description: 'Terminal value discounted at target IRR to get pre-money value. PE/VC exit analysis.',
    bsFields: [
      { key: 'revenue', label: 'Current Revenue' },
    ],
    params: [
      { key: 'projected_revenue', label: 'Revenue at Exit (₹)', required: true },
      { key: 'terminal_revenue_multiple', label: 'Exit Revenue Multiple', required: true, default: 3, min: 1, max: 15, step: 0.5 },
      { key: 'target_return', label: 'Target IRR', type: 'percent', required: true, default: 25, min: 10, max: 60, step: 5 },
      { key: 'investment', label: 'Investment Amount (₹)' },
      { key: 'years', label: 'Hold Period (yrs)', default: 5, min: 2, max: 10, step: 1 },
      { key: 'shares', label: 'Shares Outstanding', required: true },
    ],
  },
]

// ── Helpers ────────────────────────────────────────────────────────────────────
const CATEGORIES = Array.from(new Set(MODELS.map(m => m.category)))

const fmtBs = (key: string, val: any): string => {
  if (val == null) return '—'
  if (key === 'roe' || key === 'earnings_growth' || key === 'revenue_growth') return `${(val * 100).toFixed(1)}%`
  if (typeof val === 'number' && Math.abs(val) >= 1e7) {
    const cr = val / 1e7
    if (cr >= 1e5) return `₹${(cr / 1e5).toFixed(1)}L Cr`
    if (cr >= 1e3) return `₹${(cr / 1e3).toFixed(1)}K Cr`
    return `₹${cr.toFixed(0)} Cr`
  }
  if (typeof val === 'number') return val.toFixed(2)
  return String(val)
}

const fmtNum = (v: number | null | undefined, prefix = '₹'): string => {
  if (v == null || isNaN(v)) return '—'
  const abs = Math.abs(v)
  if (abs >= 1e7) return `${prefix}${(v / 1e7).toFixed(2)} Cr`
  if (abs >= 1e5) return `${prefix}${(v / 1e5).toFixed(2)} L`
  return `${prefix}${v.toFixed(2)}`
}

const fmtPct = (v: number | null | undefined): string => {
  if (v == null || isNaN(v)) return '—'
  return `${v > 0 ? '+' : ''}${v.toFixed(2)}%`
}

// ── Component ─────────────────────────────────────────────────────────────────
export default function ValuationPage() {
  const { symbol } = useParams<{ symbol: string }>()
  const navigate = useNavigate()

  const [inputs, setInputs] = useState<any>(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState('')

  const [selectedModel, setSelectedModel] = useState<ModelDef>(MODELS[0])
  const [paramValues, setParamValues] = useState<Record<string, any>>({})
  const [bsOverrides, setBsOverrides] = useState<Record<string, string>>({})

  const [result, setResult] = useState<any>(null)
  const [running, setRunning] = useState(false)
  const [runError, setRunError] = useState('')

  // ── Load inputs ─────────────────────────────────────────────────────────────
  useEffect(() => {
    if (!symbol) return
    setLoading(true)
    setError('')
    api.get(`/valuation/${symbol}/inputs`)
      .then((r: any) => {
        setInputs(r.data)
        setLoading(false)
      })
      .catch((e: any) => {
        setError(e?.response?.data?.detail || 'Failed to load valuation inputs')
        setLoading(false)
      })
  }, [symbol])

  // ── Seed params when model or inputs change ──────────────────────────────────
  useEffect(() => {
    if (!inputs) return
    const fin = inputs.financials || {}
    const suggested = inputs.suggested || {}
    const price = inputs.current_price

    const vals: Record<string, any> = {}
    selectedModel.params.forEach(p => {
      if (p.type === 'select') {
        vals[p.key] = p.options?.[0]?.value ?? 'call'
        return
      }
      // Try to find value from financials or suggested
      const raw = fin[p.key] ?? suggested[p.key]
      if (raw != null) {
        if (p.type === 'percent') {
          // Convert decimal to percent display
          const asNum = typeof raw === 'number' ? raw : parseFloat(raw)
          if (!isNaN(asNum)) {
            vals[p.key] = asNum < 2 ? parseFloat((asNum * 100).toFixed(2)) : parseFloat(asNum.toFixed(2))
          } else {
            vals[p.key] = p.default ?? 0
          }
        } else {
          vals[p.key] = raw
        }
      } else if (p.default !== undefined) {
        vals[p.key] = p.default
      }
    })

    // Special: seed price for black_scholes strike
    if ((selectedModel.id === 'black_scholes' || selectedModel.id === 'real_options') && !vals['strike']) {
      vals['strike'] = price ?? 0
    }

    setParamValues(vals)
    setResult(null)
    setRunError('')
  }, [selectedModel, inputs])

  // ── Run model ────────────────────────────────────────────────────────────────
  const runModel = useCallback(async () => {
    if (!symbol) return
    setRunning(true)
    setRunError('')
    setResult(null)

    // Convert percent values back to decimal
    const sendParams: Record<string, any> = {}
    selectedModel.params.forEach(p => {
      const v = paramValues[p.key]
      if (v === '' || v == null) return
      if (p.type === 'percent') {
        sendParams[p.key] = parseFloat(v) / 100
      } else if (p.type === 'select') {
        sendParams[p.key] = v
      } else {
        sendParams[p.key] = parseFloat(v) || v
      }
    })

    // Apply BS overrides
    const fin = inputs?.financials || {}
    selectedModel.bsFields.forEach(bf => {
      const ov = bsOverrides[bf.key]
      if (ov !== undefined && ov !== '') {
        sendParams[bf.key] = parseFloat(ov)
      }
    })

    const backendModel = selectedModel.backendModel || selectedModel.id
    try {
      const r = await api.post(`/valuation/${symbol}/run`, { model: backendModel, params: sendParams })
      setResult(r.data)
    } catch (e: any) {
      setRunError(e?.response?.data?.detail || 'Calculation failed')
    } finally {
      setRunning(false)
    }
  }, [symbol, selectedModel, paramValues, bsOverrides, inputs])

  // ── Excel export ─────────────────────────────────────────────────────────────
  const exportExcel = () => {
    if (!result) return
    const wb = XLSX.utils.book_new()

    // Summary sheet
    const summaryData = [
      ['Symbol', symbol],
      ['Model', selectedModel.label],
      ['Current Price', inputs?.current_price],
      ['Intrinsic Value', result.intrinsic_value],
      ['Upside %', result.upside_pct],
      [],
      ...Object.entries(result)
        .filter(([k]) => !['year_details', 'sensitivity', 'symbol', 'model', 'current_price'].includes(k))
        .map(([k, v]) => [k, typeof v === 'object' ? JSON.stringify(v) : v]),
    ]
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(summaryData), 'Summary')

    // Year details
    if (result.year_details?.length) {
      const keys = Object.keys(result.year_details[0])
      const ydData = [keys, ...result.year_details.map((row: any) => keys.map(k => row[k]))]
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(ydData), 'Year Details')
    }

    // Sensitivity
    if (result.sensitivity) {
      const s = result.sensitivity
      const hdr = [s.col_label, ...s.cols.map((c: number) => `${(c * 100).toFixed(1)}%`)]
      const rows = s.values.map((row: any[], i: number) => [
        `${(s.rows[i] * 100).toFixed(1)}%`,
        ...row.map((v: any) => (v ?? ''))
      ])
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([hdr, ...rows]), 'Sensitivity')
    }

    XLSX.writeFile(wb, `${symbol}_${selectedModel.id}_valuation.xlsx`)
  }

  // ── Render ────────────────────────────────────────────────────────────────────
  if (loading) return (
    <div className="flex items-center justify-center min-h-screen bg-brand">
      <div className="text-brand-accent text-lg animate-pulse">Loading valuation inputs for {symbol}...</div>
    </div>
  )

  if (error) return (
    <div className="flex flex-col items-center justify-center min-h-screen bg-brand gap-4">
      <div className="text-red-400 text-lg">{error}</div>
      <button onClick={() => navigate(-1)} className="text-brand-accent underline text-sm">Go back</button>
    </div>
  )

  const fin = inputs?.financials || {}
  const modelsAvail = inputs?.models_available || {}

  const getBsValue = (key: string) => {
    if (key === 'price') return inputs?.current_price
    return fin[key]
  }

  return (
    <div className="min-h-screen bg-brand text-slate-200">
      {/* Header */}
      <div className="border-b border-brand-border bg-brand-card px-6 py-4 flex items-center justify-between">
        <div className="flex items-center gap-4">
          <button onClick={() => navigate(-1)} className="text-slate-400 hover:text-slate-200 text-sm">← Back</button>
          <div>
            <h1 className="text-xl font-bold text-white">{symbol} — Equity Valuation</h1>
            <p className="text-slate-400 text-sm">
              Market Price: <span className="text-brand-accent font-semibold">₹{inputs?.current_price?.toFixed(2)}</span>
              <span className="mx-2 text-slate-600">|</span>
              <span className="text-slate-500">{inputs?.sector}</span>
            </p>
          </div>
        </div>
        <div className="flex gap-3">
          {result && (
            <button
              onClick={exportExcel}
              className="px-4 py-2 rounded bg-emerald-700 hover:bg-emerald-600 text-white text-sm font-medium"
            >
              Export Excel
            </button>
          )}
          <button
            onClick={runModel}
            disabled={running}
            className="px-5 py-2 rounded bg-brand-accent hover:opacity-90 text-black text-sm font-semibold disabled:opacity-50"
          >
            {running ? 'Calculating...' : 'Calculate'}
          </button>
        </div>
      </div>

      <div className="flex h-[calc(100vh-73px)]">
        {/* ── Sidebar: Model List ──────────────────────────────────────────────── */}
        <aside className="w-72 border-r border-brand-border bg-brand-card overflow-y-auto flex-shrink-0">
          {CATEGORIES.map(cat => (
            <div key={cat}>
              <div className="px-4 pt-4 pb-1 text-xs font-semibold text-slate-500 uppercase tracking-wider">{cat}</div>
              {MODELS.filter(m => m.category === cat).map(m => {
                const avail = modelsAvail[m.id] !== false
                return (
                  <button
                    key={m.id}
                    onClick={() => { setSelectedModel(m); setResult(null) }}
                    className={`w-full text-left px-4 py-2.5 text-sm transition-colors border-l-2 ${
                      selectedModel.id === m.id
                        ? 'border-brand-accent bg-brand-accent/10 text-white'
                        : 'border-transparent hover:bg-white/5 text-slate-400 hover:text-slate-200'
                    } ${!avail ? 'opacity-40' : ''}`}
                  >
                    <div className="flex items-center justify-between">
                      <span>{m.label}</span>
                      {!avail && <span className="text-xs text-slate-600 ml-1">N/A</span>}
                    </div>
                  </button>
                )
              })}
            </div>
          ))}
        </aside>

        {/* ── Main content ─────────────────────────────────────────────────────── */}
        <main className="flex-1 overflow-y-auto p-6">
          <div className="max-w-4xl mx-auto space-y-6">
            {/* Model description */}
            <div className="bg-brand-card border border-brand-border rounded-lg p-4">
              <h2 className="text-lg font-semibold text-white mb-1">{selectedModel.label}</h2>
              <p className="text-slate-400 text-sm">{selectedModel.description}</p>
            </div>

            {/* ── Section A: Balance Sheet / Financials ──────────────────────── */}
            {selectedModel.bsFields.length > 0 && (
              <div className="bg-brand-card border border-brand-border rounded-lg p-5">
                <div className="flex items-center gap-2 mb-4">
                  <span className="text-base">📋</span>
                  <h3 className="text-sm font-semibold text-slate-300">Balance Sheet / Financials</h3>
                  <span className="text-xs text-slate-500 ml-1">(from latest filing — click to override)</span>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                  {selectedModel.bsFields.map(bf => {
                    const filingVal = getBsValue(bf.key)
                    const hasOverride = bsOverrides[bf.key] !== undefined && bsOverrides[bf.key] !== ''
                    return (
                      <div key={bf.key} className="bg-brand rounded-md p-3 border border-brand-border">
                        <div className="text-xs text-slate-500 mb-1">{bf.label}</div>
                        <div className="flex items-center gap-2">
                          <span className={`text-sm font-medium ${hasOverride ? 'text-yellow-400' : 'text-slate-200'}`}>
                            {hasOverride ? fmtBs(bf.key, parseFloat(bsOverrides[bf.key])) : fmtBs(bf.key, filingVal)}
                          </span>
                        </div>
                        <input
                          type="number"
                          placeholder="Override..."
                          value={bsOverrides[bf.key] ?? ''}
                          onChange={e => setBsOverrides(prev => ({ ...prev, [bf.key]: e.target.value }))}
                          className="mt-2 w-full text-xs bg-brand-card border border-brand-border rounded px-2 py-1 text-slate-300 placeholder-slate-600 focus:outline-none focus:border-brand-accent"
                        />
                        {hasOverride && (
                          <button
                            onClick={() => setBsOverrides(prev => { const n = { ...prev }; delete n[bf.key]; return n })}
                            className="text-xs text-slate-500 hover:text-slate-300 mt-1"
                          >
                            Reset to filing
                          </button>
                        )}
                      </div>
                    )
                  })}
                </div>
              </div>
            )}

            {/* ── Section B: Assumptions & Inputs ───────────────────────────── */}
            {selectedModel.params.length > 0 && (
              <div className="bg-brand-card border border-brand-border rounded-lg p-5">
                <div className="flex items-center gap-2 mb-4">
                  <span className="text-base">⚙️</span>
                  <h3 className="text-sm font-semibold text-slate-300">Assumptions & Inputs</h3>
                  <span className="text-xs text-slate-500 ml-1">
                    (<span className="text-red-400">*</span> required)
                  </span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-5">
                  {selectedModel.params.map(p => (
                    <div key={p.key}>
                      <label className="block text-xs text-slate-400 mb-1">
                        {p.label}
                        {p.required && <span className="text-red-400 ml-1">*</span>}
                        {p.hint && <span className="text-slate-600 ml-1 font-normal">— {p.hint}</span>}
                      </label>

                      {p.type === 'select' ? (
                        <select
                          value={paramValues[p.key] ?? p.options?.[0]?.value}
                          onChange={e => setParamValues(prev => ({ ...prev, [p.key]: e.target.value }))}
                          className="w-full bg-brand border border-brand-border rounded px-3 py-2 text-sm text-slate-200 focus:outline-none focus:border-brand-accent"
                        >
                          {p.options?.map(o => (
                            <option key={o.value} value={o.value}>{o.label}</option>
                          ))}
                        </select>
                      ) : p.type === 'percent' ? (
                        <div className="space-y-1">
                          <div className="flex items-center gap-3">
                            <input
                              type="range"
                              min={p.min ?? 0}
                              max={p.max ?? 100}
                              step={p.step ?? 0.5}
                              value={paramValues[p.key] ?? p.default ?? 0}
                              onChange={e => setParamValues(prev => ({ ...prev, [p.key]: parseFloat(e.target.value) }))}
                              className="flex-1 accent-brand-accent"
                            />
                            <div className="flex items-center gap-1">
                              <input
                                type="number"
                                min={p.min}
                                max={p.max}
                                step={p.step ?? 0.1}
                                value={paramValues[p.key] ?? p.default ?? 0}
                                onChange={e => setParamValues(prev => ({ ...prev, [p.key]: parseFloat(e.target.value) || 0 }))}
                                className="w-20 bg-brand border border-brand-border rounded px-2 py-1.5 text-sm text-slate-200 text-right focus:outline-none focus:border-brand-accent"
                              />
                              <span className="text-slate-500 text-sm">%</span>
                            </div>
                          </div>
                        </div>
                      ) : (
                        <input
                          type="number"
                          step={p.step}
                          min={p.min}
                          max={p.max}
                          value={paramValues[p.key] ?? ''}
                          placeholder={p.default !== undefined ? String(p.default) : ''}
                          onChange={e => setParamValues(prev => ({ ...prev, [p.key]: e.target.value }))}
                          className="w-full bg-brand border border-brand-border rounded px-3 py-2 text-sm text-slate-200 focus:outline-none focus:border-brand-accent"
                        />
                      )}
                    </div>
                  ))}
                </div>
              </div>
            )}

            {/* trading_comps notice */}
            {selectedModel.id === 'trading_comps' && selectedModel.params.length === 0 && (
              <div className="bg-brand-card border border-brand-border rounded-lg p-5 text-sm text-slate-400">
                Trading comparables are auto-calculated using sector peers from the universe. Click <strong className="text-white">Calculate</strong> to run.
              </div>
            )}

            {/* Run error */}
            {runError && (
              <div className="bg-red-950 border border-red-700 rounded-lg p-4 text-red-300 text-sm">{runError}</div>
            )}

            {/* ── Results ────────────────────────────────────────────────────── */}
            {result && !result.error && (
              <div className="space-y-4">
                {/* Hero card */}
                <div className="bg-brand-card border border-brand-border rounded-lg p-6">
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-6">
                    <div>
                      <div className="text-xs text-slate-500 mb-1">Intrinsic Value</div>
                      <div className="text-2xl font-bold text-brand-accent">
                        {result.intrinsic_value != null ? `₹${Number(result.intrinsic_value).toFixed(2)}` : '—'}
                      </div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500 mb-1">Current Price</div>
                      <div className="text-2xl font-bold text-white">₹{inputs?.current_price?.toFixed(2)}</div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500 mb-1">Upside / Downside</div>
                      <div className={`text-2xl font-bold ${
                        result.upside_pct == null ? 'text-slate-400'
                        : result.upside_pct > 0 ? 'text-emerald-400' : 'text-red-400'
                      }`}>
                        {result.upside_pct != null ? fmtPct(result.upside_pct) : '—'}
                      </div>
                    </div>
                    <div>
                      <div className="text-xs text-slate-500 mb-1">Model</div>
                      <div className="text-sm font-semibold text-slate-200">{selectedModel.label}</div>
                    </div>
                  </div>
                </div>

                {/* Key metrics */}
                <div className="bg-brand-card border border-brand-border rounded-lg p-5">
                  <h4 className="text-sm font-semibold text-slate-300 mb-3">Model Output Details</h4>
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                    {Object.entries(result)
                      .filter(([k]) => !['year_details', 'sensitivity', 'symbol', 'model', 'current_price', 'intrinsic_value', 'upside_pct', 'recovery_breakdown', 'implied_values', 'peer_medians', 'stock_multiples'].includes(k))
                      .map(([k, v]) => {
                        if (typeof v === 'object') return null
                        const label = k.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase())
                        const display = typeof v === 'number'
                          ? (k.includes('pct') || k.includes('rate') || k.includes('irr') || k.includes('roe') || k.includes('wacc')
                              ? `${Number(v).toFixed(2)}%`
                              : Math.abs(Number(v)) >= 1e7
                                ? fmtNum(Number(v))
                                : Number(v).toFixed(2))
                          : String(v)
                        return (
                          <div key={k} className="bg-brand rounded-md p-3 border border-brand-border">
                            <div className="text-xs text-slate-500 mb-1">{label}</div>
                            <div className="text-sm font-semibold text-slate-200">{display}</div>
                          </div>
                        )
                      })}
                  </div>

                  {/* Trading comps peer medians */}
                  {result.peer_medians && (
                    <div className="mt-4">
                      <div className="text-xs text-slate-500 mb-2">Peer Median Multiples ({result.peer_count} peers)</div>
                      <div className="grid grid-cols-4 gap-3">
                        {Object.entries(result.peer_medians).map(([k, v]) => (
                          <div key={k} className="bg-brand rounded-md p-3 border border-brand-border text-center">
                            <div className="text-xs text-slate-500">{k.toUpperCase()}</div>
                            <div className="text-sm font-semibold text-slate-200 mt-1">{v != null ? Number(v).toFixed(2) : '—'}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Implied values from comps */}
                  {result.implied_values && Object.keys(result.implied_values).length > 0 && (
                    <div className="mt-4">
                      <div className="text-xs text-slate-500 mb-2">Implied Values per Share</div>
                      <div className="grid grid-cols-3 gap-3">
                        {Object.entries(result.implied_values).map(([k, v]) => (
                          <div key={k} className="bg-brand rounded-md p-3 border border-brand-border text-center">
                            <div className="text-xs text-slate-500">{k.toUpperCase()} method</div>
                            <div className="text-sm font-semibold text-brand-accent mt-1">₹{Number(v).toFixed(2)}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}

                  {/* Recovery breakdown for liquidation */}
                  {result.recovery_breakdown && (
                    <div className="mt-4">
                      <div className="text-xs text-slate-500 mb-2">Asset Recovery Breakdown</div>
                      <div className="grid grid-cols-3 sm:grid-cols-5 gap-2">
                        {Object.entries(result.recovery_breakdown).map(([k, v]) => (
                          <div key={k} className="bg-brand rounded-md p-3 border border-brand-border text-center">
                            <div className="text-xs text-slate-500 capitalize">{k}</div>
                            <div className="text-sm font-semibold text-slate-200 mt-1">{fmtNum(Number(v))}</div>
                          </div>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                {/* Year-by-year table */}
                {result.year_details && result.year_details.length > 0 && (
                  <div className="bg-brand-card border border-brand-border rounded-lg p-5">
                    <h4 className="text-sm font-semibold text-slate-300 mb-3">Year-by-Year Projection</h4>
                    <div className="overflow-x-auto">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className="border-b border-brand-border">
                            {Object.keys(result.year_details[0]).map(k => (
                              <th key={k} className="text-left py-2 pr-4 text-xs text-slate-500 font-medium">
                                {k.replace(/_/g, ' ').replace(/\b\w/g, c => c.toUpperCase())}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {result.year_details.map((row: any, i: number) => (
                            <tr key={i} className="border-b border-brand-border/50 hover:bg-white/5">
                              {Object.entries(row).map(([k, v]) => (
                                <td key={k} className="py-2 pr-4 text-slate-300">
                                  {typeof v === 'number'
                                    ? (k === 'year' || k === 'stage' ? v : (Math.abs(v) >= 1e7 ? fmtNum(v) : Number(v).toFixed(2)))
                                    : String(v ?? '—')}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {/* Sensitivity table */}
                {result.sensitivity && (
                  <div className="bg-brand-card border border-brand-border rounded-lg p-5">
                    <h4 className="text-sm font-semibold text-slate-300 mb-1">Sensitivity Analysis</h4>
                    <p className="text-xs text-slate-500 mb-3">
                      Rows: {result.sensitivity.row_label} | Columns: {result.sensitivity.col_label}
                    </p>
                    <div className="overflow-x-auto">
                      <table className="text-xs">
                        <thead>
                          <tr>
                            <th className="pr-3 py-1 text-slate-500">{result.sensitivity.row_label} ↓ / {result.sensitivity.col_label} →</th>
                            {result.sensitivity.cols.map((c: number) => (
                              <th key={c} className="pr-3 py-1 text-slate-400 font-medium">{(c * 100).toFixed(1)}%</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {result.sensitivity.values.map((row: any[], i: number) => (
                            <tr key={i} className="border-t border-brand-border/30">
                              <td className="pr-3 py-1.5 text-slate-400 font-medium">{(result.sensitivity.rows[i] * 100).toFixed(1)}%</td>
                              {row.map((val: any, j: number) => {
                                const isBase = i === 2 && j === 2
                                const price = inputs?.current_price
                                const color = val == null ? 'text-slate-600'
                                  : val > (price || 0) * 1.2 ? 'text-emerald-400'
                                  : val > (price || 0) ? 'text-emerald-600'
                                  : val > (price || 0) * 0.8 ? 'text-yellow-500'
                                  : 'text-red-400'
                                return (
                                  <td key={j} className={`pr-3 py-1.5 font-mono ${color} ${isBase ? 'ring-1 ring-brand-accent rounded px-1' : ''}`}>
                                    {val != null ? `₹${Number(val).toFixed(0)}` : '—'}
                                  </td>
                                )
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>
            )}

            {/* Result error */}
            {result?.error && (
              <div className="bg-red-950 border border-red-700 rounded-lg p-4 text-red-300 text-sm">
                Model error: {result.error}
              </div>
            )}
          </div>
        </main>
      </div>
    </div>
  )
}
