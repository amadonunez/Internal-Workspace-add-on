function getSampleCards() {
  return [buildNewCard(), buildNewCard()]
}

// Example placeholder functions (implement your actual logic)
function organizeEmails(e) {
  // Logic to organize emails
  const a = 10+10
}

function scheduleMeeting(e) {
  // Logic to schedule meetings
}

function createTask(e) {
  // Logic to create tasks
}



function buildNewCard(){

  const loremIpsum = "Lorem ipsum dolor sit amet, consec proident, sunt in culpa qui officia deserunt mollit anim id est laborum.";

  var aCardHeader = CardService.newCardHeader()
    .setImageUrl("https://ssl.gstatic.com/s2/profiles/images/silhouette200.png")
    .setImageStyle(CardService.ImageStyle.CIRCLE)
    .setTitle("Title")
    .setSubtitle("Some subtitle texts")

  const iconImage1 = CardService.newIconImage().setMaterialIcon(
    CardService.newMaterialIcon()
        .setName('mail').setWeight(500)
    );

  var aCardSection = CardService.newCardSection()
    .addWidget(CardService.newDecoratedText()
        .setStartIcon(iconImage1)
        .setTopLabel("TOP LABEL TEXT")
        .setWrapText(true)
        .setText(loremIpsum)
        .setBottomLabel("Bottom label text")
        
      )
         


  return CardService.newCardBuilder()
    .setHeader(aCardHeader)
    .addSection(aCardSection)
    .build();
}