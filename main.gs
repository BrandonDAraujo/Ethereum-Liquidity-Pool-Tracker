function myFunction() {
  const mainSheet = "main";
  const dataSheet = "data";
  const status = 1;
  const row = 3;
  const asset = 1;
  const assetPrice = 2;
  const pool = 4;
  const protocol = 3;
  const apy = 6;
  const mainS = SpreadsheetApp.getActiveSpreadsheet();
  const spread = mainS.getActiveSheet();
  SpreadsheetApp.setActiveSheet(mainS.getSheetByName("main"))
  const sRange = `A${status}:H${status}`
  mainS.getActiveSheet().getRange(sRange).setBackground("orange")

  const websites = [
    {"name": "curve", "url": 'https://curvemarketcap.com/datatable.json' },
    {"name": "badger", "url": "https://api.badger.finance/v2/setts?chain=eth" },
    {"name": "yearn", "url": "https://vaults.finance/all" },
    {"name": "appstorm", "url": "https://api.stormx.io/staking/rate" },
    {"name": "harvest_finance", "url": "https://api.harvest.finance/cmc?key=fc8ad696-7905-4daa-a552-129ede248e33" },
    {"name": "analytics_sushi", "url": "https://analytics.sushi.com/bar" },
    {"name": "vesper", "url": "https://api.vesper.finance/pools?stages=prod"},
    {"name": "nerve"},
    {"name": "ribbon", "url": {"T-WBTC":{"url": "https://api.airtable.com/v0/app5c70grFW2INfkN/T-WBTC-C?sort%5B0%5D%5Bfield%5D=Timestamp&sort%5B0%5D%5Bdirection%5D=desc&maxRecords=1"}, "T-ETH" : {"url": "https://api.airtable.com/v0/app5c70grFW2INfkN/T-ETH-C?sort%5B0%5D%5Bfield%5D=Timestamp&sort%5B0%5D%5Bdirection%5D=desc&maxRecords=1"}, "T-USDC-P-ETH": {"url": "https://api.airtable.com/v0/app5c70grFW2INfkN/T-USDC-P-ETH?sort%5B0%5D%5Bfield%5D=Timestamp&sort%5B0%5D%5Bdirection%5D=desc&maxRecords=1"}}},
    {"name" : "convex", "url": "https://www.convexfinance.com/api/curve-apys"}
    ]
  const compiledList = {};

  for (let x = 0; x < websites.length; x++) {
    let response;
    let regtest;
    let header;
    let options;
    switch (websites[x]["name"]) {
      case "curve":
        response = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        compiledList["curve"] = {}
        for (const b in response) {
          compiledList["curve"][response[b]["coin"]["label"]] = (parseFloat(response[b]["apyDay"]["label"].toFixed(4)) * 100 + (response[b]["reward"]["incentive_low"] == null ? 0 : response[b]["reward"]["incentive_low"])) / 100
        }
        break;
      case "badger":
        compiledList["badger"] = {}
        response = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        for (const b in response) {
          const a = response[b]["asset"].toUpperCase();
          compiledList["badger"][a] = response[b]["apr"] / 100
        }
        break;
      case "harvest_finance":
        compiledList["harvest_finance"] = {}
        response = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        for (let a = 0; a < response["pools"].length; a++) {
          compiledList["harvest_finance"][(response["pools"][a]["pair"]).toUpperCase()] = parseFloat(response["pools"][a]["apr"])
        }
        break;
      case "yearn":
        compiledList["yearn"] = {}
        response = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        for (let a = 0; a < response.length; a++) {
          if (response[a]["apy"] != null && response[a]["apy"]["data"]["netApy"] != null) {
            compiledList["yearn"][`${(response[a]["displayName"]).toUpperCase()}${response[a]["type"]}`] = response[a]["apy"]["data"]["netApy"]
          }
        }
        break;
      case "appstorm":
        compiledList["appstorm"] = {}
        response = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        compiledList["appstorm"]["STMX"] = response["rate"]
        break;
      case "analytics_sushi":
        compiledList["analytics_sushi"] = {}
        response = UrlFetchApp.fetch(websites[x]["url"]).getContentText();
        regtest = (/(?<=(<div class="jss39"><h6 class="MuiTypography-root MuiTypography-h6 MuiTypography-colorTextPrimary MuiTypography-noWrap">)).+?(?=(<\/h6>))/)
        compiledList["analytics_sushi"]["XSUSHI"] = response.match(regtest)[0]
        break;
      case 'vesper':
        compiledList["vesper"] = {}
        polishedResponse = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        for(const x in polishedResponse){
          compiledList["vesper"][polishedResponse[x]["name"]] = (polishedResponse[x]["earningRates"][30] + polishedResponse[x]["vspDeltaRates"][30]) / 100
        }
        break;
      case 'nerve':
        checkEmissions()
        compiledList["nerve"] = {}
        header = {"Authorization": "Basic OmMyNDVhNGEyLTYxZjgtMTFlYi1hZTkzLTAyNDJhYzEzMDAwMg=="}
        options = {"headers": header}
        const emissions = PropertiesService.getScriptProperties().getProperty("EMISSIONS")
        const nervePrice = JSON.parse(errorCheck('https://api.coingecko.com/api/v3/simple/price?ids=nerve-finance&vs_currencies=usd'))['nerve-finance']['usd']
        const tvlRaw = JSON.parse(UrlFetchApp.fetch('https://api.defistation.io/chart/Nerve%20Finance?days=30', options))
        let tvl;
        for(const x in tvlRaw["result"]){
          tvl = tvlRaw["result"][x]
        }
        const apr = ((28800 * emissions * 365) * nervePrice * 0.6) / tvl
        compiledList["nerve"][tvlRaw["defiName"]] = apr

        break;
      case 'ribbon':
        compiledList["ribbon"] = {}
        header = {"Authorization": "Bearer keymgnfgwnQHmH4pl"}
        options = {"headers": header}
        for(const a in websites[x]["url"]){
          response = JSON.parse(UrlFetchApp.fetch(websites[x]["url"][a]["url"], options).getContentText());
          compiledList["ribbon"][a] = response["records"][0]["fields"]["APY"]
        }
        break;
      case 'convex':
        compiledList["convex"] = {}
        polishedResponse = JSON.parse(UrlFetchApp.fetch(websites[x]["url"]).getContentText());
        for(const a in polishedResponse["apys"]){
          compiledList["convex"][a] = (parseFloat(polishedResponse["apys"][a]["baseApy"]) + polishedResponse["apys"][a]["crvApy"]) / 100
        }

      break;
    }
  }
    function checkEmissions(){
      const sp = PropertiesService.getScriptProperties()
      if(sp.getProperty("TIME") != null){
        let time = sp.getProperty("TIME")
        let emissions = sp.getProperty("EMISSIONS")
        const timePassed = parseInt(parseInt((Date.now() - time) / 86400000) / 7)
        if(timePassed >= 1){
          time = time + (timePassed * 604800000)
          emissions = emissions * Math.pow(0.90, timePassed)
          sp.setProperties({
            "TIME" : time,
            "EMISSIONS" : emissions
          })
        }
      }else{
        const boilerTime = 1617631200000
        const boilerEmissions = 16
        sp.setProperties({
          'TIME': boilerTime,
          'EMISSIONS': boilerEmissions
        })
        return checkEmissions()
      }
    }
    function fillAssetPrice(response){
      SpreadsheetApp.setActiveSheet(mainS.getSheetByName("main"));
      for(let c=0; c<100; c++){
        let gFormula = "="
      for(const x in response){
        gFormula += `,if(indirect(address(${row + c}, ${asset}))="${response[x]["symbol"]}", ${response[x]["current_price"]}`
      }
      gFormula += `,if(indirect(address(${row + c}, ${asset}))="", ""`
      gFormula = gFormula + ")".repeat(response.length+1);
      gFormula = gFormula.replace(/,/, "");
      mainS.getActiveSheet().getRange(row+c, assetPrice).setFormula(gFormula)
    }
    }

    function errorCheck(url){
      let options = {muteHttpExceptions: true};
      Utilities.sleep(600);
      let b = UrlFetchApp.fetch(url, options);
      let a = b.getContentText();
      if(a[0] != "{" && a[0] != "["){
        //Logger.log("Hello")
        return errorCheck(url);
      }else{
        return b;
      }
    }

  //Fill data sheet
  function formData() {
    SpreadsheetApp.setActiveSheet(mainS.getSheetByName("data"));
    mainS.getActiveSheet().clear();
    let countC = 0;
    let countR = 0;
    let maxC = [];
    let maxC1 = [];
    for (const x in compiledList) {
      mainS.getActiveSheet().getRange(1, countC + 3).setValue(x);
      for (const y in compiledList[x]) {
        maxC.push(y);
        mainS.getActiveSheet().getRange(countR + 2, countC + 3).setValue(y);
        countR++;
      }
      countR = 0;
      countC++;
      if (maxC.length > maxC1.length) {
        maxC1 = maxC;
      }
      maxC = [];
      countR = 0;
    }
    countC = 0;
    for (let x = 0; x < 100; x++) {
      mainS.getActiveSheet().getRange((x * (maxC1.length + 1)) + 1, 1).setFormula(`=INDIRECT(ADDRESS(1, MATCH(INDIRECT("main!"&ADDRESS(${row + x},${protocol})), INDIRECT(ADDRESS(1,3)&":"&ADDRESS(1,${maxC1.length + 1})), 0) + 2)&":"&ADDRESS(${maxC1.length + 1}, MATCH(INDIRECT("main!"&ADDRESS(${row + x},${protocol})), INDIRECT(ADDRESS(1,3)&":"&ADDRESS(1,${maxC1.length + 1})), 0) + 2))`)
    }
    return maxC1;
  }
  //getCoins function
  function getCoins(){
    const request = JSON.parse(errorCheck('https://api.coingecko.com/api/v3/coins/markets?vs_currency=usd&per_page=200&page=1&sparkline=false'));
    let coinList = [];
    for(const x in request){
      coinList.push(request[x]["symbol"])
    }
    return [coinList, request];
  }
  //Fill Apy function
  function fillApy() {
    SpreadsheetApp.setActiveSheet(mainS.getSheetByName("main"));
    for (let c = 0; c < 100; c++) {
      let countR = 0;
      let countC = 0;
      let gFormula = "=";
      for (const x in compiledList) {
        gFormula += `,if(INDIRECT(ADDRESS(${row + c}, ${protocol}))= "${x}"`
        for (const y in compiledList[x]) {
          gFormula += `,if(INDIRECT(ADDRESS(${row + c}, ${pool}))="${y}", ${compiledList[x][y]}`
          countC++;
        }
        gFormula += `,if(INDIRECT(ADDRESS(${row + c}, ${pool}))="", ""`
        gFormula = gFormula + ")".repeat(countC+1)
        countC = 0;
        countR++;
      }
      gFormula += `,if(INDIRECT(ADDRESS(${row + c}, ${protocol}))="", ""`
      gFormula = gFormula + ")".repeat(countR+1)
      gFormula = gFormula.replace(/,/, "")
      mainS.getActiveSheet().getRange(row + c, apy).setFormula(gFormula);
    }
  }
  //Updating
  if (mainS.getActiveSheet().getRange(row, protocol).getValue() != "") {
    const maxC1 = formData();
    fillApy();
    let sites = []
    for (const x in compiledList) {
      sites.push(x)
    }
    const everything = getCoins()
    let coinList = everything[0];
    const coinRequest = everything[1]
    fillAssetPrice(coinRequest);
    for (let c = 0; c < 100; c++) {
      mainS.getActiveSheet().getRange(row + c, asset).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(coinList).build())
      mainS.getActiveSheet().getRange(row + c, protocol).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sites).build());
      mainS.getActiveSheet().getRange(row + c, pool).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(mainS.getActiveSheet().getRange(`${dataSheet}!A${(c * (maxC1.length + 1)) + 2}:A${(c + 1) * (maxC1.length + 1)}`)).build())
    }
    //First Run
  } else {
    const everything = getCoins()
    let coinList = everything[0];
    const coinRequest = everything[1]
    fillAssetPrice(coinRequest);
    const maxC1 = formData();
    fillApy();
    let sites = [];
    let first
    for (const x in compiledList) {
      sites.push(x)
    }
    for (const x in compiledList) {
      for (const y in compiledList[x]) {
        first = y
        break
      }
      break
    }
    SpreadsheetApp.setActiveSheet(mainS.getSheetByName("main"));
    for (let c = 0; c < 100; c++) {
      mainS.getActiveSheet().getRange(row + c, protocol).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(sites).build());
      mainS.getActiveSheet().getRange(row + c, asset).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(coinList).build())
      //mainS.getActiveSheet().getRange(row + c, protocol).setValue(sites[0]);
      //mainS.getActiveSheet().getRange(row + c, pool).setValue(first);
      mainS.getActiveSheet().getRange(row + c, pool).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(mainS.getActiveSheet().getRange(`${dataSheet}!A${(c * (maxC1.length + 1)) + 2}:A${(c + 1) * (maxC1.length + 1)}`)).build())
    }


  }

  //Logger.log(compiledList)
  SpreadsheetApp.setActiveSheet(mainS.getSheetByName("main"));
  mainS.getActiveSheet().getRange(sRange).setBackground("green")
}
