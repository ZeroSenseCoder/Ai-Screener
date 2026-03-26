export const formatCurrency = (n: number | undefined | null, decimals = 2): string => {
  if (n == null) return '—'
  if (n >= 1e7) return `₹${(n / 1e7).toFixed(2)}Cr`
  if (n >= 1e5) return `₹${(n / 1e5).toFixed(2)}L`
  return `₹${n.toFixed(decimals)}`
}

export const formatMarketCap = (n: number | undefined | null): string => {
  if (n == null) return '—'
  if (n >= 1e12) return `₹${(n / 1e12).toFixed(2)}T`
  if (n >= 1e9) return `₹${(n / 1e9).toFixed(2)}B`
  if (n >= 1e7) return `₹${(n / 1e7).toFixed(2)}Cr`
  return `₹${n.toFixed(0)}`
}

export const formatPct = (n: number | undefined | null): string => {
  if (n == null) return '—'
  return `${n >= 0 ? '+' : ''}${n.toFixed(2)}%`
}

export const formatNum = (n: number | undefined | null, decimals = 2): string => {
  if (n == null) return '—'
  return n.toFixed(decimals)
}

export const formatVolume = (n: number | undefined | null): string => {
  if (n == null) return '—'
  if (n >= 1e7) return `${(n / 1e7).toFixed(2)}Cr`
  if (n >= 1e5) return `${(n / 1e5).toFixed(2)}L`
  if (n >= 1e3) return `${(n / 1e3).toFixed(1)}K`
  return `${n}`
}

export const pctColor = (n: number | undefined | null): string => {
  if (n == null) return 'text-slate-400'
  return n >= 0 ? 'text-up' : 'text-down'
}
