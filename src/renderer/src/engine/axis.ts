/**
 * Positions along one grid axis with a default size and sparse overrides (custom widths,
 * hidden rows). pos() and indexAt() are O(log k) where k = number of overrides.
 */
export class Axis {
  private idx: number[] = []
  private prefix: number[] = [0]

  constructor(
    readonly def: number,
    private sizes: Record<number, number>,
    private hidden: Record<number, true> = {},
    readonly count: number
  ) {
    const keys = new Set<number>()
    for (const k of Object.keys(sizes)) keys.add(Number(k))
    for (const k of Object.keys(hidden)) keys.add(Number(k))
    this.idx = [...keys].filter((k) => k < count).sort((a, b) => a - b)
    let acc = 0
    for (const i of this.idx) {
      acc += this.size(i) - def
      this.prefix.push(acc)
    }
  }

  size(i: number): number {
    if (this.hidden[i]) return 0
    return this.sizes[i] ?? this.def
  }

  /** number of overrides with index < i */
  private below(i: number): number {
    let lo = 0
    let hi = this.idx.length
    while (lo < hi) {
      const mid = (lo + hi) >> 1
      if (this.idx[mid] < i) lo = mid + 1
      else hi = mid
    }
    return lo
  }

  pos(i: number): number {
    return i * this.def + this.prefix[this.below(i)]
  }

  total(): number {
    return this.pos(this.count)
  }

  /** Index of the cell covering pixel offset px (clamped to the axis). */
  indexAt(px: number): number {
    if (px <= 0) return this.firstVisible(0)
    let lo = 0
    let hi = this.count - 1
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1
      if (this.pos(mid) <= px) lo = mid
      else hi = mid - 1
    }
    return lo
  }

  firstVisible(from: number): number {
    let i = from
    while (i < this.count - 1 && this.size(i) === 0) i++
    return i
  }

  /** Next visible index in a direction (skips hidden rows). */
  step(from: number, dir: 1 | -1): number {
    let i = from + dir
    while (i > 0 && i < this.count - 1 && this.size(i) === 0) i += dir
    return Math.max(0, Math.min(this.count - 1, i))
  }
}
