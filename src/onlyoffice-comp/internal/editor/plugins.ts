export const pluginsBase = "https://office-plugins.ziziyi.com/v9/sdkjs-plugins";

export const allPlugins = [
  "ai",
  "apertium",
  "autocomplete",
  "bergamot",
  "chess",
  "cvbuilder",
  "datepicker",
  "deepl",
  "doc2md",
  "drawio",
  "easybib",
  "glavred",
  "grammalecte",
  "highlightcode",
  "html",
  "icons",
  "idphoto",
  "insertQR",
  "jitsi",
  "languagetool",
  "marketplace",
  "mathpix",
  "mendeley",
  "news",
  "ocr",
  "onlydraw",
  "photoeditor",
  "pixabay",
  "pomodoro",
  "rainbow",
  "speech",
  "speechrecognition",
  "telegram",
  "termef",
  "textcleaner",
  "texthighlighter",
  "thesaurus",
  "translator",
  "typograf",
  "videoembedder",
  "wordpress",
  "wordscounter",
  "youtube",
  "zhipu",
  "zoom",
  "zotero",
];

export const featuredPlugins = [
  "marketplace",
  "ai",
  "youtube",
  "jitsi",
  "photoeditor",
  "typograf",
  "languagetool",
  "thesaurus",
  "deepl",
  "zhipu",
];

export function getPluginConfigUrl(name: string) {
  return `${pluginsBase}/${name}/config.json`;
}

export function getPluginsData(list: string[]) {
  return {
    url: "",
    pluginsData: list.map(getPluginConfigUrl),
    autostart: [],
  };
}
