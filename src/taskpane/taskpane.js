/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.OneNote) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});
const prefix = "onenote";
const urlSuffix = ".one";
const urlPrefix = "https://o.shangao.tech";

function getPageInfo(clientUrl) {
  const basePath = clientUrl.substring(prefix.length + 1, clientUrl.indexOf(urlSuffix) + urlSuffix.length);
  const params = clientUrl.substring(clientUrl.indexOf(urlSuffix) + urlSuffix.length).split("&");
  let info = {
    basePath: basePath,
    pageTitle: params[0],
  };
  params.forEach((e) => {
    // console.log(e)
    if (e.indexOf("=") !== -1) {
      let ex = e.split("=");
      info[ex[0]] = ex[1];
    }
  });
  return info;
}
const headTags = ["h1", "h2", "h3", "h4", "h5", "h6"];
function isHeadTag(html) {
  return headTags.includes(html.substring(1, 3).toLowerCase());
}

function getParagraphId(pid) {
  return pid.substring(1, pid.indexOf("}"));
}

function getLinkId(richTextId) {
  return parseInt(richTextId.substring(richTextId.indexOf("{", 1) + 1, richTextId.length - 1));
}
export async function run() {
  /**
   * Insert your OneNote code here
   */
  try {
    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
      page.load("clientUrl,webUrl");
      var outline = context.application.getActiveOutlineOrNull();

      outline.load("id, type, paragraphs/id, paragraphs/type");

      await context.sync();
      if (page.isNull) {
        return;
      }

      const pageInfo = getPageInfo(page.clientUrl);

      // console.log(pageInfo);
      if (!outline.isNull) {
        const richTextParagraphs = [];

        outline.paragraphs.items.forEach((paragraph) => {
          if (paragraph.type == "RichText") {
            const html = paragraph.richText.getHtml();
            richTextParagraphs.push([paragraph, html]);
            paragraph.load("richtext/id, richtext/text");
          }
        });

        await context.sync();
        let newContents = "<p>目录</p>";
        const newOutline = page.addOutline(200, 50, newContents);
        richTextParagraphs.forEach((p) => {
          const paragraph = p[0];
          const html = p[1];
          if (isHeadTag(html.value)) {
            const link = `${urlPrefix}onenote:${pageInfo.pageTitle}&section-id={${pageInfo[
              "section-id"
            ].toUpperCase()}}&page-id={${pageInfo["page-id"].toUpperCase()}}&object-id={${getParagraphId(
              paragraph.id
            ).toUpperCase()}}&${getLinkId(paragraph.richText.id).toString(16).toUpperCase()}&base-path=${
              pageInfo.basePath
            }`;
            const item = `<p><a href="${encodeURI(link)}">${paragraph.richText.text}</a></p>`;
            console.log("href: " + link);
            // newContents += link;
            newOutline.appendHtml(item);
            // console.log(paragraph.richText.id);
          }

          // console.log(html.value);
        });
      }
    });
  } catch (error) {
    // handler error
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  }
}
