/** Clinic Label Helper — Config (v0.1)
 *
 * IDs & constants for the Cat Vaccine Clinic label app.
 */

const CFG = {
  APP_NAME: 'Clinic Label Helper',
  // TODO: update to the actual Service Tracking Hub URL
  HUB_URL: 'https://script.google.com/macros/s/your-hub-webapp-id/exec',

  // Sheets
  SHEET_ID: '1EONCmzJHMi_1R_hfnJHAKRvXlsHsO3xSpz7cLvXjqcE',
  SHEET_NAME: 'Sheet 1',

  // Templates
  OWNER_TEMPLATE_ID: '1q6NgwBKj-6rxbih5LH57x-sLi24kX5LQ6KobfvNA2MM',
  PET_TEMPLATE_ID:   '1lmHuNjsChzOq69M0gJ9bJHWJ6Xw7DMHZxyxwwCFtfD4',

  // Output folder for merged PDFs
  OUTPUT_FOLDER_ID: '1slGMcP9iSmCWRBS2-qh-etQTQd76CpeS',

  // Merge service
  MERGE_SERVICE_URL: 'https://pdf-merge-service.onrender.com/merge', // base URL provided; using /merge

  // Column names
  COLS: {
    FIRST_NAME: 'First Name',
    LAST_NAME: 'Last Name',
    ADDRESS1: 'Address1',
    ADDRESS2: 'Address2',
    CITY: 'City',
    STATE: 'State',
    ZIP: 'Zipcode',
    PHONE: 'Phone',
    EMAIL: 'Email',
    PET_NAME: 'Pet Name',
    SPECIES: 'Type of Pet',
    AGE: 'Pet Age',
    BREED: 'Breed',
    COLOR: 'Pet Color',
    SEX: 'Pet Sex',
    SN: 'S/N',        // Spayed/Neutered? (Yes/No/Unknown)
    REASON: 'Reason', // Assumed placeholder {{Reason}}
    EVENT_DATE: 'EventDate',
    EVENT_LOCATION: 'EventLocation',
  },
};


/** Simple server-side include for HTML Service */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}