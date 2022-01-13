/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

/*
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
*/

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//#region GLOBAL VARIABLES AND REFERENCES -------------------------------------------------------------------------

  //#region IMAGE REFERENCES --------------------------------------------------------------------------------------
  // import { ContextReplacementPlugin } from "webpack";
  import "../../assets/icon-16.png";
  import "../../assets/icon-32.png";
  import "../../assets/icon-80.png";
  //#endregion ----------------------------------------------------------------------------------------------------

  //#region GLOBAL VARIABLES --------------------------------------------------------------------------------------

    //var artistColumn = "S";
    var moveEvent;
    var sortEvent;
    var sortColumn = "Priority";
    //var projectTypeColumn = "H";
    //var productColumn = "G";
   // var addedColumn = "J";
    var loop = true;
    //var startOverrideColumn = "U";
    //var workOverrideColumn = "V";
    var destinationTable;

    var brandNewBuild;
    var newBuildOtherNatives;
    var brandNewBuildFromTemplate;
    var changesToExistingNatives;
    var specCheck;
    var weTransferUpload;
    var specialRequest;
    var otherProjectType;

    var menu;
    var menuXL;
    var smallMenu;
    var brochure;
    var brochureXL;
    var smallBrochure;
    var postcard;
    var jumboPostcard;
    var colossalPostcard;
    var scratchoffPostcard;
    var jumboScratchoffPostcard;
    var peelBoxPostcard;
    var magnet;
    var foldedMagnet;
    var twoSBT;
    var boxTopper;
    var flyer;
    var doorHanger;
    var smallPlastic;
    var mediumPlastic;
    var largePlastic;
    var couponBooklet;
    var envelopeMailer;
    var birthdayPostcard;
    var newMover;
    var plasticNewMover;
    var birthdayPlastic;
    var wideFormat;
    var windowClings;
    var businessCards;
    var artworkOnly;
    var logoCreation;
    var logoRecreation;
    var legalLetter;
    var letter;
    var mapCreation;
    var menuXXL;
    var biFoldMenu;
    var mediaKit;
    var popBanner;
    var otherProduct;


    //#region TURN AROUND TIME VARIABLES ---------------------------------------------------------------------------

      var startTurnAroundTime = {
        menu: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        menuXL: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        smallMenu: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        brochure: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        brochureXL: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        smallBrochure: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        postcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        jumboPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        colossalPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        scratchoffPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        jumboScratchoffPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        peelBoxPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        magnet: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        foldedMagnet: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        twoSBT: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        boxTopper: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        flyer: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        doorHanger: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        smallPlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        mediumPlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        largePlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        couponBooklet: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        envelopeMailer: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        birthdayPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        newMover: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        plasticNewMover: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        birthdayPlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        wideFormat: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        windowClings: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        businessCards: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        artworkOnly: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        logoCreation: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        logoRecreation: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        legalLetter: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        letter: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        mapCreation: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        menuXXL: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        biFoldMenu: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        mediaKit: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        popBanner: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        }
      };

      var artTurnAroundTime = {
        menu: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        menuXL: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        smallMenu: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        brochure: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        brochureXL: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        smallBrochure: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        postcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        jumboPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        colossalPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        scratchoffPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        jumboScratchoffPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        peelBoxPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        magnet: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        foldedMagnet: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        twoSBT: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        boxTopper: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        flyer: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        doorHanger: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        smallPlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        mediumPlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        largePlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        couponBooklet: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        envelopeMailer: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        birthdayPostcard: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        newMover: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        plasticNewMover: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        birthdayPlastic: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        wideFormat: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        windowClings: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        businessCards: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        artworkOnly: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        logoCreation: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        logoRecreation: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        legalLetter: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        letter: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        mapCreation: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        menuXXL: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        biFoldMenu: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        mediaKit: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        },
        popBanner: {
          brandNewBuild: 0,
          brandNewBuildFromNatives: 0,
          brandNewBuildFromTemplate: 0,
          changesToExisitingNatives: 0,
          specCheck: 0,
          weTransferUpload: 0,
          specialRequest: 0,
          other: 0,
        }
      };

      var creativeReviewTime = {
        menu: 0,
        menuXL: 0,
        smallMenu: 0,
        brochure: 0,
        brochureXL: 0,
        smallBrochure: 0,
        postcard: 0,
        jumboPostcard: 0,
        colossalPostcard: 0,
        scratchoffPostcard: 0,
        jumboScratchoffPostcard: 0,
        peelBoxPostcard: 0,
        magnet: 0,
        foldedMagnet: 0,
        twoSBT: 0,
        boxTopper: 0,
        flyer: 0,
        doorHanger: 0,
        smallPlastic: 0,
        mediumPlastic: 0,
        largePlastic: 0,
        couponBooklet: 0,
        envelopeMailer: 0,
        birthdayPostcard: 0,
        newMover: 0,
        plasticNewMover: 0,
        birthdayPlastic: 0,
        wideFormat: 0,
        windowClings: 0,
        businessCards: 0,
        artworkOnly: 0,
        logoCreation: 0,
        logoRecreation: 0,
        legalLetter: 0,
        letter: 0,
        mapCreation: 0,
        menuXXL: 0,
        biFoldMenu: 0,
        mediaKit: 0,
        popBanner: 0,
      };
    
    //#endregion --------------------------------------------------------------------------------------------------------------


    //#region WEEKDAY VARIABLES ----------------------------------------------------------------------------------------------------------------

      var sunday = {
        dayID: 0,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }
      var monday = {
        dayID: 1,
        startHour: 8,
        startMinute: 0,
        endHour: 17,
        endMinute: 0,
        workDay: 0,
      }
      var tuesday = {
        dayID: 2,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }
      var wednesday = {
        dayID: 3,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }
      var thursday = {
        dayID: 4,
        startHour: 8,
        startMinute: 0,
        endHour: 18,
        endMinute: 0,
        workDay: 0,
      }
      var friday = {
        dayID: 5,
        startHour: 8,
        startMinute: 30,
        endHour: 13,
        endMinute: 30,
        workDay: 0,
      }
      var saturday = {
        dayID: 6,
        startHour: 8,
        startMinute: 30,
        endHour: 17,
        endMinute: 30,
        workDay: 0,
      }

      var weekdayList = [sunday, monday, tuesday, wednesday, thursday, friday, saturday];

    //#endregion --------------------------------------------------------------------------------------------------------------------------------

  //#endregion ----------------------------------------------------------------------------------------------

//#endregion ------------------------------------------------------------------------------------------------------



//#region BUTTON BEHAVIOR -----------------------------------------------------------------------------------------

  //#region CHECKBOX SETUP ________________________________________________________________________________________
  /** When the checkbox is CHANGED */
  $("#set-behavior").on("change", function() {

    // Is this set to checked?
    var checked = $(this).prop("checked");

    if (checked == true) { // Set the startup behavior!
      Office.addin.setStartupBehavior(Office.StartupBehavior.load); //when document opens, references startup behavioir in manifest, which automatically opens the taskpane
    } else { // Turn off the startup behavior!
      Office.addin.setStartupBehavior(Office.StartupBehavior.none); //when document opens, references startup behavioir in manifest, which automatically opens the taskpane
    }
  })
  //#endregion ----------------------------------------------------------------------------------------------------

//#endregion -------------------------------------------------------------------------------------------------------



//#region WATCHING FOR CHANGES ------------------------------------------------------------------------------------

  //#region STARTUP BEHAVIOR --------------------------------------------------------------------------------------
  Office.onReady((info) => {
    console.log("Office is ready!")
    // Load on Startup
    // setStartupBehavior is **document level**
    /*
    var currentBehavior = Office.addin.getStartupBehavior().then(function(returned) {
      if (returned == "Load") {
        //Check the checkbox
        $("#set-behavior").prop("checked", true);
      } else {
        //Uncheck the checkbox
        $("#set-behavior").prop("checked", false);
      }
      console.log(returned);
    });
    */
      if (info.host === Office.HostType.Excel) { //If application is Excel
        document.getElementById("sideload-msg").style.display = "none"; //Don't show side-loading message
        document.getElementById("app-body").style.display = "flex"; //Keep content in taskpane flexible to scaling, I think...
          
        Excel.run(async context => { //Do while Excel is running

          moveEvent = context.workbook.tables.onChanged.add(onTableChanged);

          // sortEvent = context.workbook.tables.onChanged.add(sortDate);

          return context.sync().then(function() { //Commits changes to document and then returns the console.log
            // console.log("Event handlers have been successfully registered");
          });
        });
      };
  });
  //#endregion ------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------



//#region MOVING AND UPDATING DATA --------------------------------------------------------------------------------

  async function onTableChanged(eventArgs) { //This function will be using event arguments to collect data from the workbook
  // async function onTableChanged(eventArgs: Excel.TableChangedEventArgs) { //TypeScript version of this command

    await Excel.run(async (context) => {      

      //#region EVENT VARIABLES -----------------------------------------------------------------------------------

        //#region EVENT ARGS --------------------------------------------------------------------------------------

          var details = eventArgs.details; //Loads the values before and after the event
          var address = eventArgs.address; //Loads the cell's address where the event took place
          var changeType = eventArgs.changeType;
          //var regexStr = address.match(/[a-zA-Z]+|[0-9]+(?:\.[0-9]+|)/g); //Separates the column letter(s) from the row number for the address: presented as a string

        //#endregion ----------------------------------------------------------------------------------------------

        //#region WORKSHEET LEVEL VARIABLES -----------------------------------------------------------------------

          var allWorksheets = context.workbook.worksheets;
          allWorksheets.load("items/name");
          //var sheet = context.workbook.worksheets.getActiveWorksheet().load("name");
          var changedWorksheet = context.workbook.worksheets.getItem(eventArgs.worksheetId).load("name");
          var worksheetTables = changedWorksheet.tables.load("items/name");
          //var changedRowSheetLevel = Number(regexStr[1]) - 1; //this variable should be used when making calculations with the changed row variable on a worksheet level (minus 1 to account for the fact that the address ignores the 0 index)
          //var changedColumnLetter = regexStr[0]; //The first instance of the separated address array, being the column letter(s)
          var changedAddress = changedWorksheet.getRange(address);
          changedAddress.load("columnIndex");
          changedAddress.load("rowIndex");

        //#endregion ----------------------------------------------------------------------------------------------

        //#region TABLE LEVEL VARIABLES ---------------------------------------------------------------------------

          var allTables = context.workbook.tables;
          allTables.load("items/name");
          var changedTable = context.workbook.tables.getItem(eventArgs.tableId).load("name"); //Returns tableId of the table where the event occured
          var startOfTable = changedTable.getRange().load("columnIndex");
          var changedTableColumns = changedTable.columns
          changedTableColumns.load("items/name");
          var changedTableRows = changedTable.rows;
          changedTableRows.load("items");
          //var changedRow = Number(regexStr[1]) - 2; //The second instance of the separated address array, being the row, converted into a number and subtracted by 2
          //it is subtracted by 2 in order to be used on a table level, which augments the row number by 2 places due to being 0 indexed and skipping the header row
          //var myRow = changedTable.rows.getItemAt(changedRow).load("values"); //loads the values of the changed row in the table where the event was fired 
          var tableRange = changedTable.getRange();

        //#endregion ----------------------------------------------------------------------------------------------

      //#endregion ------------------------------------------------------------------------------------------------

      //#region SPECIFIC TABLE VARIABLES --------------------------------------------------------------------------

        var unassignedTable = context.workbook.tables.getItem("UnassignedProjects").load("worksheet");

        var mattTable = context.workbook.tables.getItem("MattProjects").load("worksheet");

        var alainaTable = context.workbook.tables.getItem("AlainaProjects").load("worksheet");

        var bertoTable = context.workbook.tables.getItem("BertoProjects").load("worksheet");

        var breBTable = context.workbook.tables.getItem("BreBProjects").load("worksheet");

        var christianTable = context.workbook.tables.getItem("ChristianProjects").load("worksheet");

        var emilyTable = context.workbook.tables.getItem("EmilyProjects").load("worksheet");

        var ianTable = context.workbook.tables.getItem("IanProjects").load("worksheet");

        var jeffTable = context.workbook.tables.getItem("JeffProjects").load("worksheet");

        var joshTable = context.workbook.tables.getItem("JoshProjects").load("worksheet");

        var kristenTable = context.workbook.tables.getItem("KristenProjects").load("worksheet");

        var robinTable = context.workbook.tables.getItem("RobinProjects").load("worksheet");

        var lukeTable = context.workbook.tables.getItem("LukeProjects").load("worksheet");

        var lisaTable = context.workbook.tables.getItem("LisaProjects").load("worksheet");

        var luisTable = context.workbook.tables.getItem("LuisProjects").load("worksheet");

        var peterTable = context.workbook.tables.getItem("PeterProjects").load("worksheet");

        var ritaTable = context.workbook.tables.getItem("RitaProjects").load("worksheet");

        var ethanTable = context.workbook.tables.getItem("EthanProjects").load("worksheet");

        var breZTable = context.workbook.tables.getItem("BreZProjects").load("worksheet");

        var joeTable = context.workbook.tables.getItem("JoeProjects").load("worksheet");

        var jordanTable = context.workbook.tables.getItem("JordanProjects").load("worksheet");

        var hazelTable = context.workbook.tables.getItem("HazelProjects").load("worksheet");

        var toddTable = context.workbook.tables.getItem("ToddProjects").load("worksheet");

        //#region VALIDATION VARIABLES ------------------------------------------------------

          var validationSheet = context.workbook.worksheets.getItem("Validation");

          //#region PICK UP TURN AROUND TIME TABLE VARIABLES -------------------------------------
            var pickupTurnaroundTimeTable = context.workbook.tables.getItem("PickupTurnaroundTime");
            var pickupTurnaroundTimeTableRows = pickupTurnaroundTimeTable.rows;
            pickupTurnaroundTimeTableRows.load("items");
          //#endregion ---------------------------------------------------------------------------

          //#region ART TURN AROUND TIME TABLE VARIABLES -----------------------------------------
            var artTurnaroundTimeTable = context.workbook.tables.getItem("ArtTurnaroundTime");
            var artTurnaroundTimeTableRows = artTurnaroundTimeTable.rows;
            artTurnaroundTimeTableRows.load("items");
          //#endregion ---------------------------------------------------------------------------

          //#region CREATIVE REVIEW PROCESS TABLE VARIABLES --------------------------------------
            var creativeProofTable = context.workbook.tables.getItem("CreativeProofAdjust");
            var creativeProofTableRows = creativeProofTable.rows;
            creativeProofTableRows.load("items");
          //#endregion ---------------------------------------------------------------------------

          //#region PRODUCT TABLE VARIABLES (CUT MAYBE?) -----------------------------------------
            var productTable = context.workbook.tables.getItem("ProductTable");
            var productTableHoursColumn = productTable.columns.getItem("Product Hours");
            productTableHoursColumn.load("name");
            var productTableRows = productTable.rows
            productTableRows.load("items");
          //#endregion --------------------------------------------------------------------------
          
        //#endregion ----------------------------------------------------------------------------

      //#endregion -------------------------------------------------------------------------------------------------

      await context.sync().then(function () { //loads variable values

        if (changeType == "RowDeleted") {
          return;
        };

        //#region LOADED VARIABLES -------------------------------------------------------------------------------

          //#region WORKSHEET LEVEL COLUMN AND ROW INDEX NUMBERS -------------------------------------------------

            var changedColumnIndexOG = changedAddress.columnIndex;
            var changedRowIndex = changedAddress.rowIndex;

          //#endregion -------------------------------------------------------------------------------------------  
        
          //#region TABLE ITEMS AND VALUES -----------------------------------------------------------------------

            var tableColumns = changedTableColumns.items;
            var tableRows = changedTableRows.items;
            var changedRowTableIndex = changedRowIndex - 1; //adjusts index number for table level (-1 to skip header row)
            var myRow = changedTableRows.getItemAt(changedRowTableIndex); //loads the changed row in the changed table as an object
            //var oldRowValues = myRow.values;
            var rowValues = tableRows[changedRowTableIndex].values; //loads the values of the changed row in the changed table

          //#endregion -------------------------------------------------------------------------------------------

          //#region ADJUST COLUMN INDEX TO WORK ON A TABLE LEVEL -------------------------------------------------

            var tableStart = startOfTable.columnIndex;
            var changedColumnIndex = changedColumnIndexOG - tableStart;

          //#endregion -------------------------------------------------------------------------------------------

        //#endregion ---------------------------------------------------------------------------------------------

        //#region ASSIGN VALUES TO CODE FROM EXCEL --------------------------------------------------------------

          //#region ASSIGN START TURNAROUND TIME VALUES ----------------------------------------------------------

            var i = 0;
            for (var key of Object.keys(startTurnAroundTime)) { //loops through startTurnAroundTime's keys (first level objects, so menu, menuXL, postcard, etc.)
              var pickupTurnaroundTimeValues = pickupTurnaroundTimeTableRows.items[i].values; //returns values of first level object based on positon i (so if i=0, this is the menu objects. If i=1, this is menuXL objects, etc.)
              //console.log(pickupTurnaroundTimeValues[0][1]);
              startTurnAroundTime[key].brandNewBuild = pickupTurnaroundTimeValues[0][1]; //assigns brandNewBuild property of [i] sub-object the value in the first data cell in the table 
              startTurnAroundTime[key].brandNewBuildFromNatives = pickupTurnaroundTimeValues[0][2]; //assigns brandNewBuildFromNatives property of [i] sub-object the value in the second data cell in the table 
              startTurnAroundTime[key].brandNewBuildFromTemplate = pickupTurnaroundTimeValues[0][3]; //you get the point...
              startTurnAroundTime[key].changesToExisitingNatives = pickupTurnaroundTimeValues[0][4];
              startTurnAroundTime[key].specCheck = pickupTurnaroundTimeValues[0][5];
              startTurnAroundTime[key].weTransferUpload = pickupTurnaroundTimeValues[0][6];
              startTurnAroundTime[key].specialRequest = pickupTurnaroundTimeValues[0][7];
              startTurnAroundTime[key].other = pickupTurnaroundTimeValues[0][8];
              i++; //i increases so that this continues to loop through all the products, until the key gets to the end
            };

        //console.log(startTurnAroundTime);

      //#endregion --------------------------------------------------------------------------------------------

          //#region ASSIGN ART TURNAROUND TIME VALUES -------------------------------------------------------------

            var j = 0;
            for (var key of Object.keys(artTurnAroundTime)) { //loops through artTurnAroundTime's keys (first level objects, so menu, menuXL, postcard, etc.)
              var artTurnaroundTimeValues = artTurnaroundTimeTableRows.items[j].values; //returns values of first level object based on positon j (so if j=0, this is the menu objects. If j=1, this is menuXL objects, etc.)
              //console.log(artTurnaroundTimeValues[0][1]);
              artTurnAroundTime[key].brandNewBuild = artTurnaroundTimeValues[0][1]; //assigns brandNewBuild property of [j] sub-object the value in the first data cell in the table 
              artTurnAroundTime[key].brandNewBuildFromNatives = artTurnaroundTimeValues[0][2]; //assigns brandNewBuildFromNatives property of [j] sub-object the value in the second data cell in the table
              artTurnAroundTime[key].brandNewBuildFromTemplate = artTurnaroundTimeValues[0][3]; //you get it, right?
              artTurnAroundTime[key].changesToExisitingNatives = artTurnaroundTimeValues[0][4];
              artTurnAroundTime[key].specCheck = artTurnaroundTimeValues[0][5];
              artTurnAroundTime[key].weTransferUpload = artTurnaroundTimeValues[0][6];
              artTurnAroundTime[key].specialRequest = artTurnaroundTimeValues[0][7];
              artTurnAroundTime[key].other = artTurnaroundTimeValues[0][8];
              j++; //j increases so that this continues to loop through all the products, until the key gets to the end
            };

        //console.log(artTurnAroundTime);

      //#endregion --------------------------------------------------------------------------------------------

          //#region ASSIGN CREATIVE REVIEW TIME VALUES ------------------------------------------------------------

            var k = 0;
            for (var key of Object.keys(creativeReviewTime)) { //loops through creativeReviewTime's keys (first level objects, so menu, menuXL, postcard, etc.)
              var creativeReviewTimeValues = creativeProofTableRows.items[k].values; //returns values of first level object based on positon k (so if k=0, this is the menu objects. If k=1, this is menuXL objects, etc.)
              //console.log(creativeReviewTimeValues[0][1]);
              creativeReviewTime[key] = creativeReviewTimeValues[0][1]; //assigns the property of [k] sub-object the value in the first data cell in the table 
              k++; //k increases so that this continues to loop through all the products, until the key gets to the end
            };

        //console.log(creativeReviewTime);

      //#endregion --------------------------------------------------------------------------------------------

        //#endregion --------------------------------------------------------------------------------------------

        //#region ON ROW INSERTED ----------------------------------------------------------------------------------- 
          
          if (changeType == "RowInserted") {

            //#region IF CHANGE WAS NOT MADE TO VALIDATION SHEET... ----------------------------------------------------

              if (changedWorksheet.id !== validationSheet.id) {

                //#region LOAD VARIABLES ---------------------------------------------------------------------------

                  //#region LOADS COLUMN INDEXS & CELL VALUES ------------------------------------------------------

                    var addedRangeValues = cellValue(tableColumns, rowValues, "Added");
                    var startRangeValues = cellValue(tableColumns, rowValues, "Start Override");
                    var workRangeValues = cellValue(tableColumns, rowValues, "Work Override");
                    var statusColumnValue = cellValue(tableColumns, rowValues, "Status");

                    var statusColumnAddress = cellAddress(tableColumns, changedRowIndex, tableStart, changedWorksheet, "Status");

                  //#endregion -------------------------------------------------------------------------------------

                  //#region FINDS IF CHANGED TABLE IS A COMPLETED TABLE OR NOT ------------------------------------------

                    var listOfCompletedTables = [];

                    allTables.items.forEach(function (table) { //for each table in the workbook...
                      if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                        listOfCompletedTables.push(table.name); //push the name of that table into an array
                      };
                    });

                    //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
                    var includesCompletedTables = listOfCompletedTables.includes(changedTable.name);

                  //#endregion ------------------------------------------------------------------------------------------

                //#endregion ------------------------------------------------------------------------------------------

                //#region AUTOFILL COLUMN VALUES WHEN DATA IS INSERTED -------------------------------------------------

                  //#region AUTOFILL ADDED COLUMN WITH CURRENT DATE/TIME ---------------------------------------------

                    if (addedRangeValues == "") {
                      var newRange = currentDate(tableColumns, changedRowIndex, tableStart, changedWorksheet);
                      //return newRange;
                    } else {
                    console.log("Inserted row already had an Added date, so the current time was not assigned");
                    };

                  //#endregion ---------------------------------------------------------------------------------------

                  //#region AUTOFILL OVERRIDE COLUMNS WITH 0 IF EMPTY ------------------------------------------------

                    if (startRangeValues == "") {
                      var startRangeAddress = cellAddress(tableColumns, changedRowIndex, tableStart, changedWorksheet, "Start Override");
                      startRangeAddress.values = [["0"]];
                    };

                    if (workRangeValues == "") {
                      var workRangeAddress = cellAddress(tableColumns, changedRowIndex, tableStart, changedWorksheet, "Work Override");
                      workRangeAddress.values = [["0"]];
                    };

                  //#endregion ---------------------------------------------------------------------------------------

                  //#region AUTOFILL STATUS COLUMN --------------------------------------------------------------------
              
                    if (changedTable.name == "UnassignedProjects") {
                      statusColumnAddress.values = [["Awaiting Artist"]];
                    } else if (changedTable.name !== "UnassignedProjects" && includesCompletedTables == false) {
                      //if (statusColumnValue == "Awaiting Artist") {
                        statusColumnAddress.values = [["Not Working"]];
                      //};
                    } else {
                      console.log("No status column values were defaulted");
                    };
                  
                  //#endregion ----------------------------------------------------------------------------------------

                //#endregion ----------------------------------------------------------------------------------------------

              };

            //#endregion -----------------------------------------------------------------------------------------------

          };

        //#endregion --------------------------------------------------------------------------------------------------
        
        //#region ON RANGE EDITED ------------------------------------------------------------------------------------

          if (changeType == "RangeEdited" && eventArgs.details !== undefined ) {
            
            //#region ESCAPES ON TABLE CHANGED FUNCTION IF VALUES ARE UNCHANGED --------------------------------------

              // If values are the same as before, ignore the moved-to table's on change event        
              if (eventArgs.details.valueAfter == eventArgs.details.valueBefore) {
                //console.log("No values have changed. Exiting move data event...")
                return;
              };

            //#endregion --------------------------------------------------------------------------------------------

            //#region IF CHANGE WAS NOT MADE TO VALIDATION SHEET... ----------------------------------------------------

              if (changedWorksheet.id !== validationSheet.id) {
              
                //#region LOAD VARIABLES -------------------------------------------------------------------------------
      
                  //#region LOAD COLUMN INDEXES & CELL VALUES-----------------------------------------------------------------------------

                    var projectTypeColumn = findColumnPosition(tableColumns, "Project Type"); //returns the array index number of the column that matches the name of the columnName variable
                    var productColumn = findColumnPosition(tableColumns, "Product"); //returns the array index number of the column that matches the name of the columnName variable
                    var addedColumn = findColumnPosition(tableColumns, "Added"); //returns the array index number of the column that matches the name of the columnName variable
                    var statusColumn = findColumnPosition(tableColumns, "Status"); //returns the array index number of the column that matches the name of the columnName variable
                    var artistColumn = findColumnPosition(tableColumns, "Artist"); //returns the array index number of the column that matches the name of the columnName variable
                    var startOverrideColumn = findColumnPosition(tableColumns, "Start Override"); //returns the array index number of the column that matches the name of the columnName variable
                    var workOverrideColumn = findColumnPosition(tableColumns, "Work Override"); //returns the array index number of the column that matches the name of the columnName variable

                    var statusCellValue = cellValue(tableColumns, rowValues, "Status");
                    var artistCellValue = cellValue(tableColumns, rowValues, "Artist");

                  //#endregion ----------------------------------------------------------------------------------------------

                  //#region FINDS IF CHANGED TABLE IS A COMPLETED TABLE OR NOT ------------------------------------------

                    var listOfCompletedTables = [];

                    allTables.items.forEach(function (table) { //for each table in the workbook...
                      if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                        listOfCompletedTables.push(table.name); //push the name of that table into an array
                      };
                    });

                    //returns true if the changedTable is a completed table from the array previously made, false if it is anything else
                    var includesCompletedTables = listOfCompletedTables.includes(changedTable.name);

                  //#endregion ------------------------------------------------------------------------------------------

                  //#region FINDS IF CHANGE WAS MADE TO THE UNASSIGNED PROJECTS TABLE OR NOT ----------------------------

                    var isUnassigned;

                    if (changedWorksheet.name == "Unassigned Projects") {
                      isUnassigned = true;
                    } else {
                      isUnassigned = false;
                    };

                  //#endregion ------------------------------------------------------------------------------------------

                  //#region ASSIGNS THE DESTINATION TABLE VALUE ---------------------------------------------------------

                    if (artistCellValue == "Unassigned" && isUnassigned == false) {
                      destinationTable = unassignedTable;
                    } else if (artistCellValue == "Matt") {
                      destinationTable = mattTable;
                    } else if (artistCellValue == "Alaina") {
                      destinationTable = alainaTable;
                    } else if (artistCellValue == "Berto") {
                      destinationTable = bertoTable;
                    } else if (artistCellValue == "Bre B.") {
                      destinationTable = breBTable;
                    } else if (artistCellValue == "Christian") {
                      destinationTable = christianTable;
                    } else if (artistCellValue == "Emily") {
                      destinationTable = emilyTable;
                    } else if (artistCellValue == "Ian") {
                      destinationTable = ianTable;
                    } else if (artistCellValue == "Jeff") {
                      destinationTable = jeffTable;
                    } else if (artistCellValue == "Josh") {
                      destinationTable = joshTable;
                    } else if (artistCellValue == "Kristen") {
                      destinationTable = kristenTable;
                    } else if (artistCellValue == "Robin") {
                      destinationTable = robinTable;
                    } else if (artistCellValue == "Luke") {
                      destinationTable = lukeTable;
                    } else if (artistCellValue == "Lisa") {
                      destinationTable = lisaTable;
                    } else if (artistCellValue == "Luis") {
                      destinationTable = luisTable;
                    } else if (artistCellValue == "Peter") {
                      destinationTable = peterTable;
                    } else if (artistCellValue == "Rita") {
                      destinationTable = ritaTable;
                    } else if (artistCellValue == "Ethan") {
                      destinationTable = ethanTable;
                    } else if (artistCellValue == "Bre Z.") {
                      destinationTable = breZTable;
                    } else if (artistCellValue == "Joe") {
                      destinationTable = joeTable;
                    } else if (artistCellValue == "Jordan") {
                      destinationTable = jordanTable;
                    } else if (artistCellValue == "Hazel-Rah") {
                      destinationTable = hazelTable;
                    } else if (artistCellValue == "Todd") {
                      destinationTable = toddTable;
                    } else {
                      destinationTable = "null"
                    };

                  //#endregion ----------------------------------------------------------------------------------------------

                //#endregion ----------------------------------------------------------------------------------------------
                          
                //#region CLEAN UP TEXT FORMATTING ----------------------------------------------------------------------

                  //changedAddress.format.font.name = "Calibri";
                  //changedAddress.format.font.size = 12;
                  //changedAddress.format.font.color = "#000000";

                  tableRange.format.font.name = "Calibri";
                  tableRange.format.font.size = 12;
                  tableRange.format.font.color = "#000000";

                //#endregion --------------------------------------------------------------------------------------------
                    
                //#region DO FUNCTIONS ----------------------------------------------------------------------------------

                  //#region ADJUSTING TURN AROUND TIME --------------------------------------------------------------------

                    if (changedColumnIndex == projectTypeColumn || changedColumnIndex == productColumn || changedColumnIndex == addedColumn || changedColumnIndex == startOverrideColumn || changedColumnIndex == workOverrideColumn) { //if updated data was in Project Type column, run the lookupStart function

                      var startAdjustmentHours = startHoursNumber(tableColumns, rowValues, startTurnAroundTime); //adds hours to turn-around time based on Project Type
                    
                      var artAdjustmentHours = workHoursNumber(tableColumns, rowValues, artTurnAroundTime); //adds hours based on Product and adds to lookupStart output
                    
                      var artAdjustForCreativeReview = addCreativeReview(tableColumns, rowValues, artAdjustmentHours, creativeReviewTime); //takes prelookupWork variable and divides by 3 if lookupStart was equal to 2. Otherwise remains the same.
                
                      var myDate = receivedAdjust(tableColumns, rowValues); //grabs values from Added column and converts into date object in EST.
                    
                      var override = startPreAdjust(tableColumns, rowValues, startAdjustmentHours, myDate); //adds manual override start hours to adjusted start time. Adjusts for office hours and weekends.
                    
                      var startedPickedUpBy = startedBy(tableColumns, changedRowIndex, tableStart, changedWorksheet, override); //Prints the value of override to the Picked Up / Started By column and formats the date in a readible format.
                
                      var workOverride = workPrePreAdjust(tableColumns, rowValues, artAdjustForCreativeReview, override); //Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
                  
                      var proofToClient = toClient(tableColumns, changedRowIndex, tableStart, changedWorksheet, workOverride); //Prints the value of workOverride to the Proof to Client column and formats the date in a readible format.

                      console.log("Turn Around time variables were updated!");

                      return;
                  
                    };

                  //#endregion ------------------------------------------------------------------------------------------

                  //#region MOVE DATA BETWEEN SHEETS ------------------------------------------------------------------------ 

                    //#region MOVE DATA BASED ON STATUS COLUMN ------------------------------------------------------------------

                      if (changedColumnIndex == statusColumn) {

                        //#region FINDS THE COMPLETED TABLE IN CHANGED WORKSHEET ----------------------------------------------

                          var completedTable;

                          worksheetTables.items.forEach(function (table) { //for each table in the changed worksheet...
                            if (table.name.includes("Completed")) { //if the table name includes the word "Completed" in it...
                              var leTable = table.name; //sets var to name of said completed table
                              completedTable = worksheetTables.getItem(leTable); //grabs said table's data from the worksheet
                            };
                          });

                        //#endregion ------------------------------------------------------------------------------------------

                        //#region MOVES DATA TO COMPLETED TABLE ----------------------------------------------------------------

                          if (statusCellValue == "Completed" && includesCompletedTables == false && isUnassigned == false) { //if status column = "Completed", the changedTable is not a Completed table, & the changedWorksheet is not UnassignedProjects, move data to changedWorksheet's completed table
                            completedTable.rows.add(null, rowValues); //Adds empty row to bottom of the completedTable, then inserts the changed values into this empty row
                            myRow.delete(); //Deletes the changed row from the original sheet
                            console.log("Data was moved to the artist's Completed Projects Table!");
                            return;
                          } else if (statusCellValue !== "Completed" && includesCompletedTables == true) { //if status column does not = "Completed" & the changedTable is a Completed table, move data back to the artist's table
                            if (destinationTable !== "null") {
                              moveData(destinationTable, rowValues, myRow, artistCellValue);
                            };
                          };

                        //#endregion ------------------------------------------------------------------------------------------

                      };

                    //#endregion ---------------------------------------------------------------------------------------------

                    //#region MOVE DATA BASED ON ARTIST COLUMN ------------------------------------------------------------------

                      if (changedColumnIndex == artistColumn) {

                        //#region MOVES DATA TO DESTINATION TABLE ----------------------------------------------------------------

                          //if (includesCompletedTables == false) {
                            if (destinationTable !== "null") {
                              if (destinationTable.worksheet.id !== changedWorksheet.id) { //if destination table is not in the same worksheet as the changedTable (prevents for unnecessary moving of data across tables in the same worksheet), do the following...
                                moveData(destinationTable, rowValues, myRow, artistCellValue);
                                //setStatus(destinationTable, unassignedTable, tableColumns, changedRowIndex, tableStart, changedWorksheet);
                              };
                            } else {
                              console.log("No artist was assigned or updated, so no data was moved.")
                              return;
                            };
                          //};

                        //#endregion ------------------------------------------------------------------------------------------

                      };

                    //#endregion -----------------------------------------------------------------------------------------------

                  //#endregion ----------------------------------------------------------------------------------------------
                
                //#endregion ----------------------------------------------------------------------------------------------

              } else {
                  console.log("Adjustments were made to the validation sheet, therefore the date variables and move functions were not triggered");
              };

            //#endregion ----------------------------------------------------------------------------------------------------

          };

        //#endregion --------------------------------------------------------------------------------------------------------

      }).catch(function (error) {
        console.log('Error: ' + error);
        if (error instanceof OfficeExtension.Error) {
            console.log('Debug info: ' + JSON.stringify(error.debugInfo));
        };
        console.log("Promise Rejected");
      });
    });
  };

//#endregion ------------------------------------------------------------------------------------------------------



//#region SORTING -------------------------------------------------------------------------------------------------

  //#region SORT BY DATE ------------------------------------------------------------------------------------------
    async function sortDate(eventArgs) { //This function will be using event arguments to collect data from the workbook
      // console.log("SORT FUNCTION FIRED!");
      // console.log(eventArgs);

      var theChange = eventArgs.changeType; //Kind of change that was made
      var theDetails = eventArgs.details;

      // console.log("args ");

      
      if (theChange == "RangeEdited" && (theDetails == undefined || theDetails.valueTypeAfter == "String")) { //&& theDetails == undefined) {
        console.log("The sorting event has been initiated!!"); //Prevents an event from being triggered when a new row is inserted into the other sheet, thus causing duplicate runs

        //#region SORTING VARIABLES ---------------------------------------------------------------------------------
        Excel.run(async context => {
          var changedTable = context.workbook.tables.getItem(eventArgs.tableId); //Returns tableId of the table where the event occured
          var tableRange = changedTable.getRange(); //Gets the range of the changed table
          var sortHeader = tableRange.find(sortColumn, {}); //Gets the range of the entire sortColumn (the "Date" column) from the changed table
          sortHeader.load("columnIndex");
          sortHeader.load("addressLocal")
          // var sortTag = ["Urgent", "Semi-Urgent", "Not Urgent", "Eventual", "Downtime"];
          // const list = [
          //   { Tag: 'Urgent'},
          //   { Tag: 'Semi-Urgent'},
          //   { Tag: 'Not Urgent'},
          //   { Tag: 'Eventual'},
          //   { Tag: 'Downtime'},
          // ]
          //#endregion --------------------------------------------------------------------------------------------------

        //#region SORTING CONDITIONS --------------------------------------------------------------------------------
          return context.sync().then(function() {
            console.log("Sync completed...Ready to sort")
            // console.log(sortHeader.addressLocal);
            // console.log(list);

            // if (sortHeader.columnIndex == 14) {
            //   list.sort((a, b) => (a.Tag < b.Tag) ? 1 : -1);
            //   console.log(list);
            // }

            tableRange.sort.apply(
              [
                { //list of conditions to sort on
                  key: sortHeader.columnIndex, //sorts based on data in Date column
                  sortOn: Excel.SortOn.value, //sorts based on cell vlaues
                  ascending: true
                  // subField: Excel.subField, //sorts based on cell vlaues
                  // subField: String(sortTag)
                }
              ],
              false, //will not impact string ordering
              true, //table has headers
              Excel.SortOrientation.rows //sorts the rows based on previous conditions
            );

            // const myArray = [1, 2, 3, 4, 5, 6];
            // let filteredArray = list.filter((x) => {
            //   return x % 2 === 0;
            // });
            

        

            // Queue a command to apply a filter on the Category column
            // var filter = changedTable.columns.getItem("Tags").filter;
            // filter.apply({
            //     filterOn: Excel.FilterOn.values,
            //     values: ["Urgent", "Semi-Urgent", "Not Urgent", "Eventual", "Downtime"]
            // });



            console.log("Sorting is completed.")
          }); 
          //#endregion --------------------------------------------------------------------------------------------------

        }).catch(tryCatch); // CATCH EXCEL.RUN
      
      }; // END IF  
    } // END SORTDATE()
  //#endregion ----------------------------------------------------------------------------------------------------

//#endregion ------------------------------------------------------------------------------------------------------



//#region AUTOFILL FUNCTIONS -------------------------------------------------------------------------------------



  //#region GET CELL ADDRESS -------------------------------------------------------------------------------------

    /**
     * Returns the address of cell in the changedRow and columnName column for use in writing variables to cells
     * @param {Array} tableColumns An array of all the columns in the changedTable
     * @param {Number} changedRowIndex The index number of the changed row on a worksheet level
     * @param {Number} tableStart The column index where the table begins
     * @param {Object} worksheet The changed worksheet
     * @param {String} columnName The name of the column to locate in the table
     * @returns Range
     */
    function cellAddress(tableColumns, changedRowIndex, tableStart, worksheet, columnName) {

      var theColumnPosition = findColumnPosition(tableColumns, columnName); //returns the array index number of the column that matches the name of the columnName variable
      theColumnPosition = theColumnPosition + tableStart; //adjusts the column position to work on a worksheet level (in case there are multiple tables in a sheet)
      var theAddress = worksheet.getCell(changedRowIndex, theColumnPosition); //gets the range of the cell located at the given row and column positions
      return theAddress;

    };

  //#endregion ----------------------------------------------------------------------------------------------------

  //#region SET STATUS COLUMN VALUE -----------------------------------------------------------------------------

    /**
     * Assigns either "Awaiting Artist" or "Not Working" as the Status column when inserted into another table
     * @param {Object} destinationTable the table that the data is being moved to
     * @param {Object} unassignedTable the unassignedTable variable
     * @param {Array} tableColumns An array of all the columns in the changedTable
     * @param {Number} changedRowIndex The index number of the changed row on a worksheet level
     * @param {Number} tableStart The column index where the table begins
     * @param {Object} worksheet The changed worksheet
     * @returns Array of new status column values
     */
    function setStatus(destinationTable, unassignedTable, tableColumns, changedRowIndex, tableStart, worksheet) {

      if (destinationTable == unassignedTable) {
        var statusCellValue = "Awaiting Artist";
      } else {
        var statusCellValue = "Not Working";
      };

      var statusAddress = cellAddress(tableColumns, changedRowIndex, tableStart, worksheet, "Status");

      statusAddress.values = [[statusCellValue]];

      return statusAddress.values;

    };

  //#endregion ------------------------------------------------------------------------------------------------


  //#region CURRENT DATE & TIME IN ADDED COLUMN -------------------------------------------------------------------

    /**
     * Inputs the current date & time into the Added column of the changed row
     * @param {Array} tableColumns An array of all the columns in the changedTable
     * @param {Number} changedRowIndex The number of the changed row (on a worksheet level)
     * @param {Number} tableStart The column index where the table begins
     * @param {Object} worksheet The changed worksheet
     * @returns Array
     */
    function currentDate(tableColumns, changedRowIndex, tableStart, worksheet) {

      var theAddress = cellAddress(tableColumns, changedRowIndex, tableStart, worksheet, "Added");

      var now = new Date();
      var toSerial = JSDateToExcelDate(now);

      theAddress.values = [[toSerial]];
      return theAddress.values;

    };

  //#endregion ----------------------------------------------------------------------------------------------------


  //#region RETRIVE CELL VALUE -------------------------------------------------------------------------

  /**
   * Returns the value of the cell where the specified column and the changedRow intersect
   * @param {Array} tableColumns a collection of all the columns in a table in an array of objects
   * @param {Number} rowValues the values of the changed row
   * @param {String} columnName the name of the column to locate in the table
   * @returns the value of the cell where the specified column and changedRow intersect
   */
    function cellValue(tableColumns, rowValues, columnName) {

      var columnPosition = findColumnPosition(tableColumns, columnName); //returns the array index number of the column that matche4s the name of the columnName variable

      var changedRowColumnValue = rowValues[0][columnPosition]; //loads the value of the cell in the columnName column and changedRow

      return changedRowColumnValue;

  };

  //#endregion -----------------------------------------------------------------------------------------


  //#region FIND COLUMN POSITION -----------------------------------------------------------------------
  /**
   * Finds the array index number of the column that matches the name of the columnName variable
   * @param {Array} tableColumns a collection of all the columns in a table in an array of objects
   * @param {String} columnName the name of the column to locate in the table
   * @returns Number
   */
  function findColumnPosition(tableColumns, columnName) {

    var l = 0;
    for (var key of Object.keys(tableColumns)) { //loops through each column item in the tableColumns array
      var columnParent = tableColumns[l]; //gives us the array that was in whatever position l represents in the tableColumns array
      var nameOfColumn = columnParent.name; //returns the name of the column in position l
      if (nameOfColumn == columnName) { //if column name is Status, then return the position number of said column in the array to be used in the future
        var output = l;
        return output;
      } else { //otherwise, keep going
        l++;
      };
    };
  };

  //#endregion -------------------------------------------------------------------------------------------


  //#region MOVE DATA FUNCTION -----------------------------------------------------------------------------

    /**
     * moves the changed row's data to the destionation table
     * @param {Object} destinationTable the table that the data is being moved to
     * @param {Array} myRow the data, values, and attributes of the changed row
     */
    function moveData(destinationTable, rowValues, myRow, artistCellValue) {
      destinationTable.rows.add(null, rowValues); //Adds empty row to bottom of the destinationTable, then inserts the changed values into this empty row
      myRow.delete(); //Deletes the changed row from the original sheet
      console.log("Data was moved to " + artistCellValue + "'s Projects Table!");
    };

  //#endregion ---------------------------------------------------------------------------------------------


  //#region PRODUCT AND PROJECT TYPE VARIABLE ASSIGN -------------------------------------------------------

    /**
     * Converts the product and project type values into coded variables and returns an array of objects containing these two properties
     * @param {Array} tableColumns An array of all the columns in the changedTable
     * @param {Array} rowValues Loads the values of the changed row
     * @returns An object with properties
     */
    function productProjectTypeAssign(tableColumns, rowValues) {

      var productInput = cellValue(tableColumns, rowValues, "Product"); //assigns to productInput the cell value in the changed row and the Product column (a nested array of values)
      var product;

      if (productInput == "Menu") {
        product = "menu";
      } else if (productInput == "MenuXL") {
        product = "menuXL";
      } else if (productInput == "Small Menu") {
        product = "smallMenu";
      } else if (productInput == "Brochure") {
        product = "brochure";
      } else if (productInput == "BrochureXL") {
        product = "brochureXL";
      } else if (productInput == "Small Brochure") {
        product = "smallBrochure";
      } else if (productInput == "Postcard") {
        product = "postcard";
      } else if (productInput == "Jumbo Postcard") {
        product = "jumboPostcard";
      } else if (productInput == "Colossal Postcard") {
        product = "colossalPostcard";
      } else if (productInput == "Scratch-Off Postcard") {
        product = "scratchoffPostcard";
      } else if (productInput == "Jumbo Scratch-Off Postcard") {
        product = "jumboScratchoffPostcard";
      } else if (productInput == "Peel-A-Box Postcard") {
        product = "peelBoxPostcard";
      } else if (productInput == "Magnet") {
        product = "magnet";
      } else if (productInput == "Folded Magnet") {
        product = "foldedMagnet";
      } else if (productInput == "2SBT") {
        product = "twoSBT";
      } else if (productInput == "Box Topper") {
        product = "boxTopper";
      } else if (productInput == "Flyer") {
        product = "flyer";
      } else if (productInput == "Door Hanger") {
        product = "doorHanger";
      } else if (productInput == "Small Plastic") {
        product = "smallPlastic";
      } else if (productInput == "Medium Plastic") {
        product = "mediumPlastic";
      } else if (productInput == "Large Plastic") {
        product = "largePlastic";
      } else if (productInput == "Coupon Booklet") {
        product = "couponBooklet";
      } else if (productInput == "Envelope Mailer") {
        product = "envelopeMailer";
      } else if (productInput == "Birthday Postcard") {
        product = "birthdayPostcard";
      } else if (productInput == "New Mover") {
        product = "newMover";
      } else if (productInput == "Plastic New Mover") {
        product = "plasticNewMover";
      } else if (productInput == "Birthday Plastic") {
        product = "birthdayPlastic";
      } else if (productInput == "Wide Format") {
        product = "wideFormat";
      } else if (productInput == "Window Clings") {
        product = "windowClings";
      } else if (productInput == "Business Cards") {
        product = "businessCards";
      } else if (productInput == "Artwork Only") {
        product = "artworkOnly";
      } else if (productInput == "Logo Creation") {
        product = "logoCreation";
      } else if (productInput == "Logo Recreation") {
        product = "logoRecreation";
      } else if (productInput == "Legal Letter") {
        product = "legalLetter";
      } else if (productInput == "Letter") {
        product = "letter";
      } else if (productInput == "Map Creation") {
        product = "mapCreation";
      } else if (productInput == "MenuXXL") {
        product = "menuXXL";
      } else if (productInput == "Bi-Fold Menu") {
        product = "biFoldMenu";
      } else if (productInput == "Media Kit") {
        product = "mediaKit";
      } else if (productInput == "POP Banner") {
        product = "popBanner";
      } else {
        product = "";
      };

      var projectTypeInput = cellValue(tableColumns, rowValues, "Project Type"); //assigns projectTypeInput the cell value in the changed row and the Project Type column (a nested array of values)

      var projectType;

      if (projectTypeInput == "Brand New Build") {
        projectType = "brandNewBuild";
      } else if (projectTypeInput == "Brand New Build from Other Product Natives") {
        projectType = "brandNewBuildFromNatives";
      } else if (projectTypeInput == "Brand New Build From Template") {
        projectType = "brandNewBuildFromTemplate";
      } else if (projectTypeInput == "Changes to Exisiting Natives") {
        projectType = "changesToExistingNatives";
      } else if (projectTypeInput == "Specification Check") {
        projectType = "specCheck";
      } else if (projectTypeInput == "WeTransfer Upload to MS") {
        projectType = "weTransferUpload";
      } else if (projectTypeInput == "Special Request") {
        projectType = "specialRequest";
      } else if (projectTypeInput == "Other") {
        projectType = "other";
      } else {
        projectType = "";
      };

      return {
        product,
        projectType
      };
    }

  //#endregion -------------------------------------------------------------------------------------------------------


  //#region START ADJUSTMENT HOURS -----------------------------------------------------------------------------------

    /**
     * Finds the value of Project Type and Product in the changed row and returns a number of hours for the Start By Turn Around Time
     * @param {Array} tableColumns An array of all the columns in the changedTable
     * @param {Array} rowValues Loads the values of the changed row
     * @param {Object} startTurnAroundTime A variable containing objects that represent all the given values for start time based on project type and product, pulled from the validation sheet
     * @returns Number
     */   
    function startHoursNumber(tableColumns, rowValues, startTurnAroundTime) { //loads these variables from another function to use in this function

      var input = productProjectTypeAssign(tableColumns, rowValues); //Converts the product and project type values into coded variables and returns an array of objects containing these two properties
      var productInput = input.product //the product value in the changed row
      var projectTypeInput = input.projectType; //the project type value in the changed row

      if (productInput == "") {
        console.log("Product column cell is empty, therefore start turn around time calucations could not be executed");
        return;
      } else if (projectTypeInput == "") {
        console.log("Project Type column cell is empty, therefore start turn around time calucations could not be executed");
        return;
      };

      var startHours = startTurnAroundTime[productInput][projectTypeInput]; //uses info from product and project type columns to retrun the proper value from the startTurnAroundTime variable

      return startHours;

    };
  
  //#endregion ---------------------------------------------------------------------------------------------------


  //#region PICKED UP / STARTED BY -------------------------------------------------------------------------------


    //#region MY DATE ----------------------------------------------------------------------------------------------

      /**
       * Finds the value of Date Added in the changed row and converts it to be a date object in EST.
       * @param {Array} tableColumns An array of all the columns in the changedTable
       * @param {Array} rowValues Loads the values of the changed row
       * @returns Date
       */
      function receivedAdjust(tableColumns, rowValues) {

        var dateTime = cellValue(tableColumns, rowValues, "Added"); //assigns to productInput the cell value in the changed row and the Product column (a nested array of values)

        //var dateTime = rowValues[0][9]; //assigns input the cell value in the changed row and the Added column (a nested array of values)

        var date = new Date(Math.round((dateTime - 25569)*86400*1000)); //convert serial number to date object
        date.setMinutes(date.getMinutes() + date.getTimezoneOffset()); //adjusting from GMT to EST (adds 4 hours)
        return date;
      };

    //#endregion ---------------------------------------------------------------------------------------------------

    //#region OVERRIDE ---------------------------------------------------------------------------------------------

      /**
       * Finds the value of Start Override in the changed row and adds it to projectTypeHours, then adds that new number as hours to myDate. Adjusts for office hours and weekends.
       * @param {Array} tableColumns An array of all the columns in the changedTable
       * @param {Array} rowValues loads the values of the changed row
       * @param {Number} startAdjustmentHours returned number representing amount of hours before project needs to be picked up
       * @param {Date} myDate receivedAdjust returned date
       * @return {Date}
       */
      function startPreAdjust(tableColumns, rowValues, startAdjustmentHours, myDate) {

        var startOverride = cellValue(tableColumns, rowValues, "Start Override"); //gets values of Start Orverride cell
        var startManualOverride = startAdjustmentHours + startOverride; //adds start override value to the start hours adjustment
        var myDateCopy = new Date(myDate); //sets myDateCopy to myDate as a new date variable (so the old date doesnt get changed)
        var adjustedDateTime = officeHours(myDateCopy, startManualOverride); //converts to be within office hours if it already isn't
        return adjustedDateTime;

      };

    //#endregion ----------------------------------------------------------------------------------------------------

    //#region STARTED PICKED UP BY ---------------------------------------------------------------------------------

      /**
       * Prints the value of override to the Picked Up / Started By column and formats the date in a readible format
       * @param {Array} tableColumns An array of all the columns in the changedTable
       * @param {Number} changedRowIndex The index number of the changed row on a worksheet level
       * @param {Number} tableStart The column index where the table begins
       * @param {Object} worksheet The changed worksheet
       * @param {Date} override date adjusted for office hours
       * @returns date
       */
      function startedBy(tableColumns, changedRowIndex, tableStart, worksheet, override) { //loads these variables from another function to use in this function

        var theAddress = cellAddress(tableColumns, changedRowIndex, tableStart, worksheet, "Picked Up / Started By");
            
        var serialDate = JSDateToExcelDate(override);  //converts override date object back into a serial number

        theAddress.values = [[serialDate]]; //assigns the returned serial number to the cell
        return theAddress.values; //commits changes and exits the function

      };

    //#endregion ----------------------------------------------------------------------------------------------------

  //#endregion ------------------------------------------------------------------------------------------------------


  //#region PROOF TO CLIENT --------------------------------------------------------------------------------------

    //#region ART ADJUSTMENT HOURS -----------------------------------------------------------------------------------
    /**
     * Finds the value of Project Type and Product in the changed row and returns a number of hours for the Art Turn Around Time
     * @param {Array} tableColumns An array of all the columns in the changedTable
     * @param {Array} rowValues Loads the values of the changed row
     * @param {Object} artTurnAroundTime A variable containing objects that represent all the given values for art working time based on project type and product, pulled from the validation sheet
     * @returns Number
     */   
     function workHoursNumber(tableColumns, rowValues, artTurnAroundTime) { //loads these variables from another function to use in this function

      var input = productProjectTypeAssign(tableColumns, rowValues); //Converts the product and project type values into coded variables and returns an array of objects containing these two properties
      var productInput = input.product //the product value in the changed row
      var projectTypeInput = input.projectType; //the project type value in the changed row

      if (productInput == "") {
        console.log("Product column cell is empty, therefore art turn around time calucations could not be executed");
        return;
      } else if (projectTypeInput == "") {
        console.log("Project Type column cell is empty, therefore art turn around time calucations could not be executed");
        return;
      };

      var workHours = artTurnAroundTime[productInput][projectTypeInput]; //uses info from product and project type columns to retrun the proper value from the artTurnAroundTime variable

      return workHours;

    };
  
  //#endregion ---------------------------------------------------------------------------------------------------

    //#region ART ADJUST FOR CREATIVE REVIEW ------------------------------------------------------------------------------------
      /**
       * if Project Type value is anything other than a new build (and friends), adjusts the Product Hours number to be a third of it's normal value, resulting in a shorter proof to client time
       * @param {Array} tableColumns An array of all the columns in the changedTable
       * @param {Number} rowValues loads the values of the changed row
       * @param {Number} artAdjustmentHours returned number representing amount of hours before proof needs to be submitted to client
       * @param {Number} creativeReviewTime a variable containing objects that represent all the given values for creative review process time based on product, pulled from the validation sheet
       * @returns Number
       */
      function addCreativeReview(tableColumns, rowValues, artAdjustmentHours, creativeReviewTime) {

        var input = productProjectTypeAssign(tableColumns, rowValues); //Converts the product and project type values into coded variables and returns an array of objects containing these two properties
        var productInput = input.product //the product value in the changed row

        if (productInput == "") {
          console.log("Product column cell is empty, therefore creative proof turn around time calucations could not be executed");
          return;
        };

        var creativeHours = creativeReviewTime[productInput]; //loads the creative review hours for the specific product

        var adjustedForCreativeReview = artAdjustmentHours + creativeHours; //adds creative review hours to art adjustment hours found in previous function

        return adjustedForCreativeReview;

      };

    //#endregion ---------------------------------------------------------------------------------------------------

    //#region WORKOVERRIDE --------------------------------------------------------------------------------------------
      /**
       * Finds the value of Work Override in the changed row and adds it to workHoursAdjust, then adds that new number as hours to startedPickedUpBy. Formats to be within office hours and on a weekday if needed.
       * @param {Array} tableColumns An array of all the columns in the changedTable
       * @param {Array} rowValues loads the values of the changed row
       * @param {Number} artAdjustForCreativeReview returned number representing amount of hours before proof needs to be submitted to client
       * @param {Date} override loads the date that the project should be picked up by
       * @returns Date
       */
      function workPrePreAdjust (tableColumns, rowValues, artAdjustForCreativeReview, override) {

        var workOverride = cellValue(tableColumns, rowValues, "Work Override"); //gets values of Work Orverride cell
        var workManualAdjust = artAdjustForCreativeReview + workOverride; //adds start override value to the number of hours for the project type
        var overrideCopy = new Date(override); //sets overrideCopy to a new date variable (so the old date doesnt get changed)
        var adjustedDateTime = officeHours(overrideCopy, workManualAdjust);
        return adjustedDateTime;

      };
    //#endregion --------------------------------------------------------------------------------------------------

    //#region PROOF TO CLIENT -----------------------------------------------------------------------------------------

      /**
       * Prints the value of workOverride to the Proof to Client column and formats the date in a readible format
       * @param {Array} tableColumns An array of all the columns in the changedTable
       * @param {Number} changedRowIndex The index number of the changed row on a worksheet level
       * @param {Number} tableStart The column index where the table begins
       * @param {Object} worksheet The changed worksheet
       * @param {Date} workOverride proof to client date found in the workPreAdjust function (after converted to be within office hours and on a weekday)
       * @returns date
       */
      function toClient(tableColumns, changedRowIndex, tableStart, worksheet, workOverride) { //loads these variables from another function to use in this function

        var theAddress = cellAddress(tableColumns, changedRowIndex, tableStart, worksheet, "Proof to Client");

        var serialDateTheSecond = JSDateToExcelDate(workOverride); //converts workOverride date object back into a serial number

        theAddress.values = [[serialDateTheSecond]]; //assigns the returned serial number to the cell
        return theAddress.values; //commits changes and exits the function

      };

    //#endregion ----------------------------------------------------------------------------------------------------

  //#endregion ------------------------------------------------------------------------------------------------------


  //#region OFFICE HOURS ---------------------------------------------------------------------------------------
    /**
     * Sets weekday variables and loops through the withinOfficeHours function, which adjusts the date to be within office hours
     * @param {Date} date Date to be adjusted to be within office hours
     * @param {Number} number Number of adjustment hours to add to date
     * @returns Date
     */
    function officeHours(day, number) {

      //#region SETTING WORKDAY HOURS IN THE WEEKDAY VARIABLES -------------------------------------------------------------------------------------

        //loops through my weekday variables, finds returns the proper variable title for it's index in the array, and then runs it through the findWorkDay function
        for (var i = 0; i < weekdayList.length; i++) {
          var weekdayReplacement = findWorkDay(weekdayList[i]);
        };

      //#endregion --------------------------------------------------------------------------------------------------------------------------------

      //var aNum = 0

      while (loop == true) {
      var officeHours = withinOfficeHours(day, number);
      day = officeHours.date;
      number = officeHours.adjustmentNumber;
      loop = officeHours.loop;
      //aNum++
      };
      //console.log("The correct date & time is: " + day);
      loop = true;
      return day;
    };

      //#region FUNCTIONS -------------------------------------------------------------------------------------------------------------------------

        //#region WITHIN OFFICE HOURS FUNCTION -------------------------------------------------------------------------------------------------
          /**
           * Adjusts date to be within office hours while maintaining an accurate turn around time variable for the adjustment number
           * @param {Date} date Date to be adjusted to be within office hours 
           * @param {Number} adjustmentNumber Number of adjustment hours to add to date
           * @returns An object with properties (date, adjustment number, and loop)
           */
          function withinOfficeHours(date, adjustmentNumber) {

            //#region VARIABLES ------------------------------------------------------------------------------------------------------------

              //#region SETS DATE VARIABLES ----------------------------------------------------------------------------------------------

                //converts our input variables into milliseconds
                var dateMilli = date.getTime();
                var adjustmentNumberMilli = adjustmentNumber * 3600000;

                //gets day of the week attributes for the date variable
                var dateDayOfWeek = dayOfWeek(date); //returns a dayID (0-6) for the day of the week of the date object
                var dayTitle = titleDOW(dateDayOfWeek); //returns a day title based on the dayID of the dateDayOfWeek variable

                //retrives workday variables associated with the weekday of the date variable
                var bookendVars = startEndMidnight(date, dayTitle);

                  //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED OUTSIDE OF OFFICE HOURS ---------------------------------------

                    if (date < bookendVars.startOfWorkDayMilli) { //if date is between 12AM and start time, adjust hours to be the start time
                        date.setHours(dayTitle.startHour);
                        date.setMinutes(dayTitle.startMinute);
                        date.setSeconds(0);
                        dateMilli = date.getTime();
                        bookendVars = startEndMidnight(date, dayTitle);
                    };

                    if (date > bookendVars.endOfWorkDayMilli) { //if date is after end time and before 12AM, go to next day and adjust hours to be the start time of that next day
                        date.setDate(date.getDate() + 1);
                        dateDayOfWeek = dayOfWeek(date);
                        dayTitle = titleDOW(dateDayOfWeek);
                        date.setHours(dayTitle.startHour);
                        date.setMinutes(dayTitle.startMinute);
                        date.setSeconds(0);
                        dateMilli = date.getTime();
                        bookendVars = startEndMidnight(date, dayTitle);
                    };
                  
                  //#endregion ------------------------------------------------------------------------------------------------------------

                  //#region ADJUSTS DATES IN CASE REQUEST WAS SUBMITTED ON WEEKEND ----------------------------------------------------

                        if ((dateDayOfWeek == 6) || (dateDayOfWeek == 0)) { //if date was submitted on a weekend...
                          date = weekendAdjust(date, dateDayOfWeek);
                          dateDayOfWeek = dayOfWeek(date);
                          dayTitle = titleDOW(dateDayOfWeek);
                          date.setHours(dayTitle.startHour);
                          date.setMinutes(dayTitle.startMinute);
                          date.setSeconds(0);
                          dateMilli = date.getTime();
                          bookendVars = startEndMidnight(date, dayTitle);
                        };
              
                      //#endregion ------------------------------------------------------------------------------------------------------------

              //#endregion ----------------------------------------------------------------------------------------------------------------

              //#region SETS ADJUSTMENT DATE VARIABLES -----------------------------------------------------------------------------------

                //adds adjustmentNumber to date to get an adjustedDate value that will be used in later checks and calculations
                var adjustedDate = new Date(date);
                var adjustedDateMilli = adjustedDate.getTime();
                adjustedDateMilli = adjustedDateMilli + adjustmentNumberMilli;
                adjustedDate = new Date(adjustedDateMilli);

              //#endregion ---------------------------------------------------------------------------------------------------------------

              //#region SETS ADD A DAY VARIABLES -----------------------------------------------------------------------------------------

                  //gets day of the week attributes for the day after the date variable
                    var nextDay = new Date(date);

                    var newNextDay = getNextDay(nextDay); //also sets this variable to the start time of the next day
                    var addADay = newNextDay.nextDay;
                    var addADayTitle = newNextDay.nextDayTitle;
                    var addADayMilli = addADay.getTime();
                    
                    //retrives workday variables associated with the weekday of the addADay variable
                    var bookendAddedDate = startEndMidnight(addADay, addADayTitle);

                //#endregion ----------------------------------------------------------------------------------------------------------------

            //#endregion --------------------------------------------------------------------------------------------------------------------

            //#region ACTION: SETS ADJUSTED DATE TO BE WITHIN OFFICE HOURS ------------------------------------------------------------------

              //if adjustedDate falls outside of office hours, do this...
              if (adjustedDateMilli < bookendVars.startOfWorkDayMilli || adjustedDateMilli > bookendVars.endOfWorkDayMilli) { //since the bookendVars is in reference to the date variable, this function will still trigger if adjustedDate is technically within office hours, but on a different day

                //#region SETS ADJUSTMENT NUMBER VALUES ---------------------------------------------------------------------------------

                  var dayRemainder = (((bookendVars.endOfWorkDayMilli - dateMilli) / 1000) / 60) / 60; //time between end of work day and the original date time
                  var remainingAdjust = adjustmentNumber - dayRemainder; //gives us the remaining adjustment hours based off of what was already used to get to the end of the work day
                  var remainingAdjustMilli = remainingAdjust * 3600000;

                //#endregion ------------------------------------------------------------------------------------------------------------

                //#region NEW DAY CALCULATIONS ------------------------------------------------------------------------------------------

                  var newDay = new Date(addADay);

                  //adds remaining adjustment hours to the beginning of the work day the next day after date (addADay)
                  var dateTimeAdjusted = newDay.setMilliseconds((newDay.getMilliseconds() + remainingAdjustMilli));

                  var dateTimeAdjustedConvert = new Date(dateTimeAdjusted); //convert serial number to date object

                  date = dateTimeAdjustedConvert; //not sure if it should be date or something else yet. Need to make sure that the function works with this

                //#endregion ------------------------------------------------------------------------------------------------------------

                //#region SET LOOP VARIABLES IF STILL NOT WITHIN OFFICE HOURS OR EXCEEDS OFFICE HOURS OF NEXT DAY -----------------------

                    //if the new date exceeds the office hours of addADay, then do this...
                    if (dateTimeAdjusted > bookendAddedDate.endOfWorkDayMilli) {
                      adjustmentNumber = (remainingAdjust - addADayTitle.workDay) //subtracts remainingAdjust hours from the total workDay hours in the addADay variable
                      var dayAfterTomorrow = new Date(addADay);
                      var newDayAfterTomorrow = getNextDay(dayAfterTomorrow);
                      date = new Date(newDayAfterTomorrow.nextDay);
                      loop = true;
                      return {
                        date,
                        adjustmentNumber,
                        loop
                      };
                    } else {
                      loop = false;
                      return {
                        date,
                        //adjustmentNumber,
                        loop
                      };
                    };

                  //#endregion -------------------------------------------------------------------------------------------------------------
              
              } else {
                date = adjustedDate;
                loop = false;
                return {
                  date,
                  adjustmentNumber,
                  loop
                };
              };
            
            //#endregion --------------------------------------------------------------------------------------------------------------------

          };

        //#endregion ---------------------------------------------------------------------------------------------------------------------------


        //#region FIND WORK DAY FUNCTION -------------------------------------------------------------------------------------------------------

          /**
           * Returns the number of hours in a specific work day by subtracting the start from the end of the day, based on the properties loaded by the weekday variable
           * @param {Object} weekday A weekday variable with all it's associated properties
           * @returns Number
           */
          function findWorkDay(weekday) {

            //sets start time for weekday variable to a date for calculations
            var start = new Date(0); //69, baby
            start.setHours(weekday.startHour);
            start.setMinutes(weekday.startMinute);
            start.setSeconds(0);

            //sets end time for weekday variable to a date for calculations
            var end = new Date(0); //seriously though, just making sure the dates for both variables will always be the same
            end.setHours(weekday.endHour);
            end.setMinutes(weekday.endMinute);
            end.setSeconds(0);

            var workDayTime = (((end.valueOf() - start.valueOf()) / 1000) / 60) / 60; //subtracts end of day from start of day to get total work day hours for that weekday, then converts the milliseconds into hours (with decimal for minutes, if any)

            weekday.workDay = workDayTime; //sets our number to the variable 

            return weekday.workDay //returns our number to the actual object variable outside of the function

          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------


        //#region DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------------

          /**
           * Returns a number 0-6 (Sunday - Saturday) based on the date input
           * @param {Date} d loads a date variable
           * @returns Number
           */
          function dayOfWeek(d) { //finds the day of the week
            var day = d.getDay();
            return day;
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region TITLE DAY OF WEEK FUNCTION ---------------------------------------------------------------------------------------------------

          /**
           * Returns the weekday variable, with all it's associated properties, from the weekday index input value
           * @param {Number} d The indexed number (0-6) of the weekday
           * @returns An object with properties
           */
          function titleDOW(d) { //returns the day of the week (refered to directly in another variable) based on the dayID index number
            if (d == 0) {
              return sunday;
            } else if (d == 1) {
              return monday;
            } else if (d == 2) {
              return tuesday;
            } else if (d == 3) {
              return wednesday;
            } else if (d == 4) {
              return thursday;
            } else if (d == 5) {
              return friday;
            } else if (d == 6) {
              return saturday;
            };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region START/END/MIDNIGHT FUNCTION --------------------------------------------------------------------------------------------------
          
          /**
           * Sets start of work day, end of work day, and midnight to a date variable (including millisecond versions), returning an object with specific properties for each
           * @param {Date} originalDate A date variable (in this case, the date before any alterations)
           * @param {object} weekday A weekday variable with all its associated properties
           * @returns An object with properties
           */
          function startEndMidnight(originalDate, weekday) {

            var startOfWorkDay = new Date(originalDate); //adjusts start time of work day based on the day of the week
            startOfWorkDay.setHours(weekday.startHour);
            startOfWorkDay.setMinutes(weekday.startMinute);
            startOfWorkDay.setSeconds(0);
            var startOfWorkDayMilli = startOfWorkDay.getTime();

            var endOfWorkDay = new Date(originalDate); //adjusts end time of work day based on the day of the week
            endOfWorkDay.setHours(weekday.endHour);
            endOfWorkDay.setMinutes(weekday.endMinute);
            endOfWorkDay.setSeconds(0);
            var endOfWorkDayMilli = endOfWorkDay.getTime();

            var midnight = new Date(originalDate);
            midnight.setDate(midnight.getDate() + 1);
            midnight.setHours(0);
            midnight.setMinutes(0);
            midnight.setSeconds(0);
            var midnightMilli = midnight.getTime();

            return {
              startOfWorkDay,
              startOfWorkDayMilli,
              endOfWorkDay,
              endOfWorkDayMilli,
              midnight,
              midnightMilli
            };

          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region GET NEXT DAY FUNCTION --------------------------------------------------------------------------------------------------------

          /**
           * Adds a day to the date variable and sets it to the start time of that new day's day of the week. Also adjusts for weekends if needed.
           * @param {Date} date A date object
           * @returns An object with properties
           */
          function getNextDay(date) {

            var nextDay = new Date(date);
            var newNextDay = nextDay.setDate(nextDay.getDate() + 1); //returns the day after the original date
            nextDay = new Date(newNextDay);
            var nextDayDayOfWeek = dayOfWeek(nextDay);
            var nextDayTitle = titleDOW(nextDayDayOfWeek); //returns a day title based on the dayID of the addADay variable

              if ((nextDayDayOfWeek == 6) || (nextDayDayOfWeek == 0)) { //checks if nextDay falls on a weekend
                nextDay = weekendAdjust(nextDay, nextDayDayOfWeek); //adjusts nextDay output to not fall on a weekend
                nextDayDayOfWeek = dayOfWeek(nextDay);
                nextDayTitle = titleDOW(nextDayDayOfWeek);
              };

              nextDay.setHours(nextDayTitle.startHour);
              nextDay.setMinutes(nextDayTitle.startMinute);
              nextDay.setSeconds(0);
              return {
                nextDay,
                nextDayTitle
              };
          };

        //#endregion ----------------------------------------------------------------------------------------------------------------------------------


        //#region WEEKEND ADJUST FUNCTION ------------------------------------------------------------------------------------------------------
          
          /**
           * If input date falls on a weekend, returns a new date adjusted to start on the next upcoming Monday
           * @param {Date} date A date variable
           * @param {Number} dateWeekday A number indexed 0-6 representing the weekday of the date variable
           * @returns Date
           */
          function weekendAdjust(date, dateWeekday) {
            if (dateWeekday == 6) {
              var weekend = new Date(date);
              weekend.setDate(weekend.getDate() + 2);
              return weekend;
            } else if (dateWeekday == 0) {
              var weekend = new Date(date);
              weekend.setDate(weekend.getDate() + 1);
              return weekend;
            };
          };

        //#endregion ------------------------------------------------------------------------------------------------------------------------------


        //#region CONVERT DATE TO SERIAL ----------------------------------------------------------------------------------------------------------

          /**
           * Converts input date into serial number that excel can apply conditional formatting to
           * @param {Date} inDate A date variable
           * @returns String
           */
          function JSDateToExcelDate(inDate) {

            var returnDateTime = 25569.0 + ((inDate.getTime() - (inDate.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24));
            //var returnDateTime = 25569.0 + ((inDate.getTime()) / (1000 * 60 * 60 * 24));
            return returnDateTime.toString().substr(0,20);
        
          };

        //#endregion --------------------------------------------------------------------------------------------------------------------------------

      //#endregion -------------------------------------------------------------------------------------------------------------------------------

  //#endregion -------------------------------------------------------------------------------------------------------------------------------------


//#endregion ---------------------------------------------------------------------------------------------------------------------------------------



//#region ERROR HANDLING ------------------------------------------------------------------------------------------

  //#region TRY CATCH ---------------------------------------------------------------------------------------------
  async function tryCatch(callback) {
    try {
      await callback();
    } catch (error) {
      console.error(error);
    }
  }
  //#endregion ---------------------------------------------------------------------------------------------------

//#endregion -----------------------------------------------------------------------------------------------------

