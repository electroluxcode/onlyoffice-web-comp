import { ONLYOFFICE_LANG_KEY } from "../const";

export type OnlyOfficeLang =
  (typeof ONLYOFFICE_LANG_KEY)[keyof typeof ONLYOFFICE_LANG_KEY];

let currentLang: OnlyOfficeLang = ONLYOFFICE_LANG_KEY.ZH;

export function getCurrentLang(): OnlyOfficeLang {
  return currentLang;
}

export function setCurrentLang(lang: OnlyOfficeLang) {
  currentLang = lang;
}

export function getOnlyOfficeLang() {
  return getCurrentLang();
}
