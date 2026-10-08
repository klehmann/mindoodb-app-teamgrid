// Strings of the view-sheet commands and dialog (Data tab, group "MindooDB").
// Classic TeamGrid's German labels where it had them; other languages fall
// back to English.
export const VIEW_SHEET_STRINGS = {
  en: {
    group: 'MindooDB view',
    add: 'View sheet',
    addTip: 'Add a sheet filled from a MindooDB view',
    refresh: 'Refresh',
    refreshTip: 'Fill the active view sheet again from its view',
    settings: 'Settings',
    settingsTip: 'Change the view, rows and name of the active view sheet',
    addTitle: 'Add virtual view sheet',
    settingsTitle: 'View sheet settings',
    nameLabel: 'Sheet name',
    viewLabel: 'View',
    showDocuments: 'Show documents',
    showCategories: 'Show categories',
    rootLabel: 'Top level category',
    rootPlaceholder: 'Category\\Subcategory',
    hint: 'A view sheet shows a virtual view configured for this app in Haven. Refreshing it rewrites the sheet.',
    noViews: 'No views are configured for this app yet. Add one as a data source in the app settings in Haven.',
    notViewSheet: 'The active sheet is not a view sheet.',
    viewGone: 'The view "{view}" is no longer configured for this app.',
    nameMissing: 'Enter a sheet name.',
    nameTaken: 'A sheet with this name already exists.',
    nothingShown: 'Show documents, categories or both.',
    failed: 'The view could not be read: {error}',
    readOnly: 'You can only read this workbook.',
    apply: 'Apply',
    close: 'Close',
  },
  de: {
    group: 'MindooDB-Ansicht',
    add: 'Ansichtsblatt',
    addTip: 'Ein Blatt mit den Daten einer MindooDB-Ansicht hinzufügen',
    refresh: 'Aktualisieren',
    refreshTip: 'Das aktive Ansichtsblatt neu aus seiner Ansicht füllen',
    settings: 'Einstellungen',
    settingsTip: 'Ansicht, Zeilen und Namen des aktiven Ansichtsblatts ändern',
    addTitle: 'Virtuelles Ansichtsblatt hinzufügen',
    settingsTitle: 'Einstellungen des Ansichtsblatts',
    nameLabel: 'Blattname',
    viewLabel: 'Ansicht',
    showDocuments: 'Dokumente anzeigen',
    showCategories: 'Kategorien anzeigen',
    rootLabel: 'Oberste Kategorie',
    rootPlaceholder: 'Kategorie\\Unterkategorie',
    hint: 'Ein Ansichtsblatt zeigt eine virtuelle Ansicht, die in Haven für diese App eingerichtet ist. Beim Aktualisieren wird das Blatt neu geschrieben.',
    noViews: 'Für diese App ist noch keine Ansicht eingerichtet. Füge in Haven in den App-Einstellungen eine als Datenquelle hinzu.',
    notViewSheet: 'Das aktive Blatt ist kein Ansichtsblatt.',
    viewGone: 'Die Ansicht „{view}“ ist für diese App nicht mehr eingerichtet.',
    nameMissing: 'Gib einen Blattnamen ein.',
    nameTaken: 'Ein Blatt mit diesem Namen gibt es schon.',
    nothingShown: 'Zeige Dokumente, Kategorien oder beides.',
    failed: 'Die Ansicht konnte nicht gelesen werden: {error}',
    readOnly: 'Diese Arbeitsmappe kannst du nur lesen.',
    apply: 'Übernehmen',
    close: 'Schließen',
  },
} satisfies Record<string, Record<string, string>>

export type ViewSheetStrings = (typeof VIEW_SHEET_STRINGS)['en']

export function viewSheetStrings(language: string): ViewSheetStrings {
  return (VIEW_SHEET_STRINGS as Record<string, ViewSheetStrings>)[language] ?? VIEW_SHEET_STRINGS.en
}
