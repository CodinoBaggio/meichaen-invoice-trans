const scriptProp = PropertiesService.getScriptProperties();

const scriptProperty = (() => ({
  awaitingSendFolder: scriptProp.getProperty("awaitingSendFolder"),
  sentFolder: scriptProp.getProperty("sentFolder"),
}))();

// const scriptProperty = (function () {
//   const scriptProp = PropertiesService.getScriptProperties();
//   return {
//     backetName: scriptProp.getProperty("folderId"),
//     privateKey: scriptProp.getProperty("private_key").replace(/\\n/g, '\n'),
//     serviceAccountEmail: scriptProp.getProperty("client_email"),
//     smaregiId: scriptProp.getProperty("smaregiId"),
//     smaregiToken: scriptProp.getProperty("smaregiToken"),
//     calendarSpreadSheetId: scriptProp.getProperty("calendarSpreadSheetId"),
//   }
// })();
