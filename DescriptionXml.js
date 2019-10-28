/*
Name: Calendar Description XML Updater
Author: Alex Wolfe
*/

function getDateUpdated(xml) {
  try {
    var desc = parseCalendarDescription(xml);
    return desc.getAttribute('updated').getValue();
  } catch(e) {
    return null;
  }
}

function getVersionUsed(xml) {
  try {
    var desc = parseCalendarDescription(xml);
    return desc.getAttribute('version').getValue();
  } catch(e) {
    return null;
  }
}

function getPersonalizationArrayfromXML(xml) {
  try {
    var desc = parseCalendarDescription(xml);
    var p = desc.getChildren('personalization');
    if (p.length!=14) {
      Logger.log("Unexpected number of personalizations in XML ("+p.length+")");
      return null;}
    var pArray = new Array(p.length);
    for (i=0; i<pArray.length; i++) {
      var x = new Array(3);
      x = [p[i].getAttribute('n').getValue(), p[i].getAttribute('c').getValue(), p[i].getAttribute('l').getValue()];
      pArray[i] = x;
    }
    return pArray;
  } catch(e) {
    return null;
  }
}

function parseCalendarDescription(xml) {
  try {
    var document = XmlService.parse(xml);
    var root = document.getRootElement();
    return root;
  } catch(e) {
    return null;
  }
}

function writeCalendarDescription(calendar, version, updated, personalizationArray) {
  var root = XmlService.createElement('bmsaschedule')
    .setAttribute('version',version)
    .setAttribute('updated',updated);
  //Logger.log("Writing Personalizations to XML");
  /*for (i=0; i<personalizationArray.length; i++) {
    var p = XmlService.createElement('p')
      .setAttribute('n', personalizationArray[i][0])
      .setAttribute('c', personalizationArray[i][1])
      .setAttribute('l', personalizationArray[i][2]);
    root.addContent(p);
  }*/
  var document = XmlService.createDocument(root);
  var xml = XmlService.getPrettyFormat().format(document);
  calendar.setDescription(xml);
}
