var pcfButtonEvent = (function () {
  let _formContext;

  const pcfControlFieldName = "tec_downloadpdfreportbutton";

  function onLoad(executionContext) {
    const formContext = executionContext.getFormContext();
    _formContext = formContext;

    const pcfControl = formContext.getControl(pcfControlFieldName);
    if (pcfControl) {
      pcfControl.addEventHandler("onButtonClick", pcfButtonClicked);
    }
  }

  function pcfButtonClicked(params) {
    alert("PCF Button Clicked!");
    alert(params.Id);
    alert(params.Name);
  }

  //public api
  return {
    onLoad: onLoad,
    pcfButtonClicked: pcfButtonClicked,
  };
})();

{
  "compilerOptions": {
    "allowJs": true
  },
  "include": [
    "*.js"
  ]
}
