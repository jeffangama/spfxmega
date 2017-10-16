Access dev environment

  go to a modern ui page,
  add this to the url

  ?loadSPFX=true&debugManifestsFile=https://localhost:4321/temp/manifests.js&customActions={"a861c815-e425-416d-9520-04bcdf557e27":{"location":"ClientSideExtension.ApplicationCustomizer"}}

  sessionStorage.clear()


How to deploy to tenant :

  1. deploy to style library the following files :
  a. file in temp deploy
  b. css

  In console, type sessionStorage.clear() to refresh the menu
