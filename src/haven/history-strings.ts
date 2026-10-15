// Strings of the revision browser and the read-only banners, taken from
// classic TeamGrid's locales (src/i18n/locales/*.json on main) so both apps
// say the same. Languages outside this list fall back to English.

export const HISTORY_STRINGS = {
  "en": {
    "browseRevisions": "Browse revisions",
    "title": "Spreadsheet revisions",
    "intro": "Pick a saved spreadsheet revision to open it read-only.",
    "loading": "Loading revisions...",
    "empty": "No revisions are available for this spreadsheet.",
    "current": "Current revision",
    "deleted": "Deleted",
    "listFailed": "The revision list could not be loaded.",
    "loadFailed": "The revision could not be loaded.",
    "historicalBanner": "You're viewing a historical revision - read-only.",
    "timeTravelBanner": "Time travel mode is active as of {date} - read-only.",
    "returnToCurrent": "Return to current"
  },
  "de": {
    "browseRevisions": "Revisionen durchsuchen",
    "title": "Tabellenrevisionen",
    "intro": "Wähle eine gespeicherte Tabellenrevision, um sie schreibgeschützt zu öffnen.",
    "loading": "Revisionen werden geladen...",
    "empty": "Für diese Tabelle sind keine Revisionen verfügbar.",
    "current": "Aktuelle Revision",
    "deleted": "Gelöscht",
    "listFailed": "Die Revisionsliste konnte nicht geladen werden.",
    "loadFailed": "Die Revision konnte nicht geladen werden.",
    "historicalBanner": "Du siehst eine historische Revision – schreibgeschützt.",
    "timeTravelBanner": "Der Zeitreisemodus ist seit {date} aktiv – schreibgeschützt.",
    "returnToCurrent": "Zur aktuellen Version"
  },
  "fr": {
    "browseRevisions": "Parcourir les révisions",
    "title": "Révisions de la feuille",
    "intro": "Choisissez une révision enregistrée pour l’ouvrir en lecture seule.",
    "loading": "Chargement des révisions...",
    "empty": "Aucune révision n’est disponible pour cette feuille.",
    "current": "Révision actuelle",
    "deleted": "Supprimée",
    "listFailed": "La liste des révisions n’a pas pu être chargée.",
    "loadFailed": "La révision n’a pas pu être chargée.",
    "historicalBanner": "Vous consultez une révision historique – lecture seule.",
    "timeTravelBanner": "Le mode voyage dans le temps est actif au {date} — lecture seule.",
    "returnToCurrent": "Revenir à la version actuelle"
  },
  "it": {
    "browseRevisions": "Sfoglia revisioni",
    "title": "Revisioni del foglio",
    "intro": "Scegli una revisione salvata per aprirla in sola lettura.",
    "loading": "Caricamento revisioni...",
    "empty": "Nessuna revisione disponibile per questo foglio.",
    "current": "Revisione corrente",
    "deleted": "Eliminata",
    "listFailed": "Impossibile caricare l’elenco delle revisioni.",
    "loadFailed": "Impossibile caricare la revisione.",
    "historicalBanner": "Stai visualizzando una revisione storica – sola lettura.",
    "timeTravelBanner": "La modalità viaggio nel tempo è attiva al {date} — sola lettura.",
    "returnToCurrent": "Torna alla corrente"
  },
  "es": {
    "browseRevisions": "Explorar revisiones",
    "title": "Revisiones de la hoja",
    "intro": "Elija una revisión guardada para abrirla en solo lectura.",
    "loading": "Cargando revisiones...",
    "empty": "No hay revisiones disponibles para esta hoja.",
    "current": "Revisión actual",
    "deleted": "Eliminada",
    "listFailed": "No se pudo cargar la lista de revisiones.",
    "loadFailed": "No se pudo cargar la revisión.",
    "historicalBanner": "Está viendo una revisión histórica: solo lectura.",
    "timeTravelBanner": "El modo de viaje en el tiempo está activo desde {date}: solo lectura.",
    "returnToCurrent": "Volver a la actual"
  },
  "nl": {
    "browseRevisions": "Revisies doorbladeren",
    "title": "Spreadsheetrevisies",
    "intro": "Kies een opgeslagen revisie om deze alleen-lezen te openen.",
    "loading": "Revisies laden...",
    "empty": "Er zijn geen revisies beschikbaar voor deze spreadsheet.",
    "current": "Huidige revisie",
    "deleted": "Verwijderd",
    "listFailed": "De revisielijst kon niet worden geladen.",
    "loadFailed": "De revisie kon niet worden geladen.",
    "historicalBanner": "U bekijkt een historische revisie – alleen-lezen.",
    "timeTravelBanner": "De tijdreismodus is actief per {date} — alleen-lezen.",
    "returnToCurrent": "Terug naar huidig"
  },
  "nb": {
    "browseRevisions": "Bla gjennom revisjoner",
    "title": "Regnearkrevisjoner",
    "intro": "Velg en lagret revisjon for å åpne den skrivebeskyttet.",
    "loading": "Laster revisjoner...",
    "empty": "Ingen revisjoner er tilgjengelige for dette regnearket.",
    "current": "Gjeldende revisjon",
    "deleted": "Slettet",
    "listFailed": "Revisjonslisten kunne ikke lastes.",
    "loadFailed": "Revisjonen kunne ikke lastes.",
    "historicalBanner": "Du ser på en historisk revisjon – skrivebeskyttet.",
    "timeTravelBanner": "Tidsreisemodus er aktiv per {date} – skrivebeskyttet.",
    "returnToCurrent": "Tilbake til gjeldende"
  },
  "pl": {
    "browseRevisions": "Przeglądaj rewizje",
    "title": "Rewizje arkusza",
    "intro": "Wybierz zapisaną rewizję, aby otworzyć ją w trybie tylko do odczytu.",
    "loading": "Ładowanie rewizji...",
    "empty": "Brak dostępnych rewizji dla tego arkusza.",
    "current": "Bieżąca rewizja",
    "deleted": "Usunięta",
    "listFailed": "Nie udało się wczytać listy rewizji.",
    "loadFailed": "Nie udało się wczytać rewizji.",
    "historicalBanner": "Przeglądasz historyczną rewizję – tylko do odczytu.",
    "timeTravelBanner": "Tryb podróży w czasie jest aktywny na dzień {date} — tylko do odczytu.",
    "returnToCurrent": "Wróć do bieżącej"
  }
} as const

export type HistoryStrings = Record<keyof (typeof HISTORY_STRINGS)['en'], string>

export function historyStrings(language: string): HistoryStrings {
  const all: Record<string, HistoryStrings> = HISTORY_STRINGS
  return all[language] ?? all.en!
}
