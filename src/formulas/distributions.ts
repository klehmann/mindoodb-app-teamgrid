/**
 * Binomial and negative binomial distributions as Excel computes them: closed forms in
 * log space and the regularized incomplete beta function, so billions of trials neither
 * loop nor underflow.
 */

/** ln Γ(x) for x > 0 (Lanczos, g = 7, n = 9; about 15 significant digits). */
export function logGamma(x: number): number {
  const c = [
    0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
    12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7,
  ]
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - logGamma(1 - x)
  const z = x - 1
  let sum = c[0]!
  for (let i = 1; i < 9; i++) sum += c[i]! / (z + i)
  const t = z + 7.5
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(sum)
}

/** ln C(n, k). */
function logChoose(n: number, k: number): number {
  return logGamma(n + 1) - logGamma(k + 1) - logGamma(n - k + 1)
}

/** Continued fraction of the incomplete beta function (modified Lentz). */
function betaContinuedFraction(a: number, b: number, x: number): number {
  const TINY = 1e-300
  let c = 1
  let d = 1 - ((a + b) * x) / (a + 1)
  if (Math.abs(d) < TINY) d = TINY
  d = 1 / d
  let h = d
  for (let m = 1; m <= 10_000; m++) {
    const m2 = 2 * m
    let aa = (m * (b - m) * x) / ((a + m2 - 1) * (a + m2))
    d = 1 + aa * d
    if (Math.abs(d) < TINY) d = TINY
    c = 1 + aa / c
    if (Math.abs(c) < TINY) c = TINY
    d = 1 / d
    h *= d * c
    aa = (-(a + m) * (a + b + m) * x) / ((a + m2) * (a + m2 + 1))
    d = 1 + aa * d
    if (Math.abs(d) < TINY) d = TINY
    c = 1 + aa / c
    if (Math.abs(c) < TINY) c = TINY
    d = 1 / d
    const delta = d * c
    h *= delta
    if (Math.abs(delta - 1) < 1e-16) break
  }
  return h
}

/** Regularized incomplete beta function I_x(a, b). */
export function regularizedBeta(x: number, a: number, b: number): number {
  if (x <= 0) return 0
  if (x >= 1) return 1
  const front = Math.exp(logGamma(a + b) - logGamma(a) - logGamma(b) + a * Math.log(x) + b * Math.log1p(-x))
  return x < (a + 1) / (a + b + 2)
    ? (front * betaContinuedFraction(a, b, x)) / a
    : 1 - (front * betaContinuedFraction(b, a, 1 - x)) / b
}

export function binomialPdf(k: number, n: number, p: number): number {
  if (p === 0) return k === 0 ? 1 : 0
  if (p === 1) return k === n ? 1 : 0
  return Math.exp(logChoose(n, k) + k * Math.log(p) + (n - k) * Math.log1p(-p))
}

/** P(X ≤ k) = I_{1−p}(n − k, k + 1). */
export function binomialCdf(k: number, n: number, p: number): number {
  if (k >= n) return 1
  if (p === 0) return 1
  if (p === 1) return 0
  // Small cases: the direct sum is exact to the last bit, as Excel's results are.
  if (k <= 1000) {
    let sum = 0
    for (let i = 0; i <= k; i++) sum += binomialPdf(i, n, p)
    return Math.min(sum, 1)
  }
  return regularizedBeta(1 - p, n - k, k + 1)
}

/** The smallest k with P(X ≤ k) ≥ alpha. */
export function binomialInverse(n: number, p: number, alpha: number): number {
  let low = 0
  let high = n
  while (low < high) {
    const middle = Math.floor((low + high) / 2)
    if (binomialCdf(middle, n, p) >= alpha) high = middle
    else low = middle + 1
  }
  return low
}

/** P(f failures before the s-th success). */
export function negativeBinomialPdf(f: number, s: number, p: number): number {
  return Math.exp(logChoose(f + s - 1, f) + s * Math.log(p) + f * Math.log1p(-p))
}

/** P(F ≤ f) = I_p(s, f + 1). */
export function negativeBinomialCdf(f: number, s: number, p: number): number {
  if (f <= 1000) {
    let sum = 0
    for (let i = 0; i <= f; i++) sum += negativeBinomialPdf(i, s, p)
    return Math.min(sum, 1)
  }
  return regularizedBeta(p, s, f + 1)
}
