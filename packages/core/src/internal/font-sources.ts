// Internal font-source descriptors, factories and substitution data. Not yet
// part of the package root API: the `fonts` load option is wired format by
// format, and the factories become public together with it.
export {
  FONT_SOURCE_PROTOCOL,
  assertFontSourceDescriptors,
  cssApi,
  cssApiRequests,
  fontFaceCss,
  fontFiles,
  fontResolver,
  fontResolverFunction,
  validateFontFaceDeclaration,
  type CssApiOptions,
  type CssApiRequest,
  type CssApiSource,
  type FontDemand,
  type FontFaceCssOptions,
  type FontFaceCssSource,
  type FontFaceDeclaration,
  type FontFilesOptions,
  type FontFilesSource,
  type FontResolveContext,
  type FontResolveFn,
  type FontResolverOptions,
  type FontResolverSource,
  type FontSource,
  type FontSourceCredentials,
} from '../fonts/source.js';
export {
  OFFICE_FONT_SUBSTITUTIONS,
  fontFamilyKey,
  loadedRegularAliases,
  matchFontDeclaration,
  substitutionsFor,
  type FontDeclarationMatch,
  type FontSubstitution,
  type LoadedFaceFacts,
} from '../fonts/substitutions.js';
export {
  GOOGLE_FONTS_CATALOG,
  GOOGLE_FONTS_DEFAULT_CSS_ORIGIN,
  GOOGLE_NOTO_CJK_NAMING,
  googleFonts,
  type GoogleFontsOptions,
} from '../fonts/presets/google-fonts.js';
export { fontSourceLimitError } from '../errors/font-source-limit.js';
export {
  FontSourceSession,
  normalizeFontDemand,
  type FontResolverRunner,
  type FontSourceSessionOptions,
  type LoadedFontSources,
  type LoadedSourceFace,
} from '../fonts/source-session.js';
export { fetchFontStylesheet, parseFontFaceCss, parseUnicodeRange } from '../fonts/css-source.js';
