export {
  createSpellcheckModule,
  replaceInParagraph,
  spellcheckRibbonItems as ribbonItems,
  SPELLCHECK_MODULE_ID,
  SPELLING_ERROR_COLOR,
} from './module';
export type { SpellcheckModule, SpellcheckModuleOptions, SpellingCommandArgs } from './module';
export { SpellChecker, localStorageSpellStorage, type Misspelling, type SpellSettings, type SpellStorage } from './checker';
export { Hunspell, matchCase, wordCase } from './hunspell';
export { findWords, type WordSpan } from './words';
export { DEFAULT_LANGUAGE, LANGUAGES, type Language } from './dictionaries';
