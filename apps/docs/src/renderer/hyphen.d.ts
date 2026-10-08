/** hyphen ships no types: a Liang pattern hyphenator and its per-language pattern tables */
declare module 'hyphen' {
  export default function createHyphenator(
    patterns: unknown,
    options?: { hyphenChar?: string; minWordLength?: number },
  ): (text: string) => string
}
declare module 'hyphen/patterns/en-us' {
  const patterns: unknown
  export default patterns
}
