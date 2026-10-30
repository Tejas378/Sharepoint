const CONFIG = {
  // lid : "5FA2A1C8-1712-4250-89F8-B253AD7539D4", //user groups list id
    // lid: "261c8455-7b94-467f-aef1-edff2186573b", //List emp
    // lid: "199b8847-a1ff-4abe-88b1-e48e5453c64e", //List Eplooyee
    // lid: "3ea14eef-d55f-421e-a42a-260cdd4fa8c6", //List emp
    // lid: "bfd7d181-1c24-49ee-a84d-c6f7d9cd4a7c", //List use list with shareid's
    // lid: "1b1776d1-b907-48a8-9d47-6ddb31968e82", //Library
    lid: "9d2bfd03-db02-429b-b000-c07aca6db30f",
    itemId: 6,
    folderId: 9,
    listName: "Invoices",
    LibraryName: "Invoices",
    subFolderName: "Invoices",
    viewName: "Invoices",
    fileName:"Attackerfile.txt",
    SITE_URL: "https://tmsobb2.sharepoint.com",
    personGroup: ["Communication site Visitors","Communication site Members","Communication site Owners","attacker@tmsobb2.onmicrosoft.com","tmsobb@tmsobb2.onmicrosoft.com","i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com","i:0#.f|membership|attacker@tmsobb2.onmicrosoft.com"],
    userLoginName : "i:0#.f|membership|tmsobb@tmsobb2.onmicrosoft.com",
    kinds: [2, 3, 4, 5, 6],
    scope: [0, 1, 2, 3],
    groupIds: [5,3,4,29,20,25,21,12,13,14,41,26,30,44,47,38,37,33,36,22,50 ],
    groupId: 50,
    groupName : "Test",
    userid:16,
    roles: [1073741829,1073741828,1073741830,1073741827,1073741826,1073741832,1073741825,1073741833,1073741924,1073741925,1073741926],
    roles2: [5, 6, 7, 8, 27, 32, 28, 30, 29, 3, 0, 1, 2, 4, 3],  //Share links role kind
    principleIds: [15,16,9,8,7,45,48,18,23,27,31,34,39,42,10,1073741822,1073741823,11],
    HTTPMethods: ["GET", "POST", "PUT", "PATCH", "MERGE", "DELETE"]
  };
  

  // DEPRECATED: this web part still keeps its own configuration and hardcoded
  // tenant URLs. It is not the active research UI; the supported path is
  // src/webparts/hacktheworld (central config + components/config.ts).
  //
  // The XHR helper below is fixed in place for the remaining doclib code:
  //   - bodies are sent as UTF-8 strings instead of char-code Blobs;
  //   - the digest is acquired from /_api/contextinfo and cached/expired;
  //   - requests time out instead of hanging;
  //   - non-2xx responses reject (they already did here).
  const READ_METHODS = ["GET", "HEAD", "OPTIONS"];
  const REQUEST_TIMEOUT_MS = 60000;
  const LEGACY_HOST = "tmsobb2.sharepoint.com";

  let cachedDigest = null;
  let pendingDigest = null;

  function rewriteLegacyHost(url) {
    try {
      const parsed = new URL(url);
      const site = new URL(CONFIG.SITE_URL);
      if (parsed.host.toLowerCase() === LEGACY_HOST && parsed.origin !== site.origin) {
        return site.origin + parsed.pathname + parsed.search + parsed.hash;
      }
      return url;
    } catch (error) {
      return url;
    }
  }

  function getDigest(url) {
    if (cachedDigest && cachedDigest.expiresAt > Date.now()) {
      return Promise.resolve(cachedDigest.value);
    }
    if (pendingDigest) {
      return pendingDigest;
    }
    const contextUrl = new URL(url).origin + "/_api/contextinfo";
    pendingDigest = fetch(contextUrl, {
      method: "POST",
      headers: { Accept: "application/json;odata=verbose" },
      credentials: "include"
    })
      .then(function (response) {
        if (!response.ok) {
          throw new Error("contextinfo failed with HTTP " + response.status);
        }
        return response.json();
      })
      .then(function (data) {
        const info = data && data.d && data.d.GetContextWebInformation;
        if (!info || !info.FormDigestValue) {
          throw new Error("contextinfo response did not contain FormDigestValue");
        }
        const seconds = Number(info.FormDigestTimeoutSeconds) || 1800;
        cachedDigest = {
          value: info.FormDigestValue,
          expiresAt: Date.now() + Math.max(0, seconds - 60) * 1000
        };
        pendingDigest = null;
        return cachedDigest.value;
      }, function (error) {
        pendingDigest = null;
        throw error;
      });
    return pendingDigest;
  }

  const sendRequest = (method, url, body = "", parseJson = true) => {
    const intended = String(method || "GET").toUpperCase();
    const isRead = READ_METHODS.indexOf(intended) >= 0;
    const target = rewriteLegacyHost(url);
    const headers = {
      Accept: "application/json;odata=verbose",
      "Accept-Language": "en-GB,en-US;q=0.9,en;q=0.8,mr;q=0.7"
    };
    let sentMethod = intended;
    if (intended === "MERGE") {
      sentMethod = "POST";
      headers["X-HTTP-Method"] = "MERGE";
    }
    const payload =
      body === "" || body === undefined || body === null
        ? undefined
        : typeof body === "string"
        ? body
        : JSON.stringify(body);

    return Promise.resolve()
      .then(function () {
        if (isRead) {
          return null;
        }
        headers["Content-Type"] = "application/json;odata=verbose";
        return getDigest(target);
      })
      .then(function (digest) {
        if (digest) {
          headers["X-RequestDigest"] = digest;
        }
        const controller = typeof AbortController !== "undefined" ? new AbortController() : null;
        const timer = controller
          ? setTimeout(function () {
              controller.abort();
            }, REQUEST_TIMEOUT_MS)
          : null;

        return fetch(target, {
          method: sentMethod,
          headers: headers,
          credentials: "include",
          body: payload,
          signal: controller ? controller.signal : undefined
        }).then(
          function (response) {
            if (timer) {
              clearTimeout(timer);
            }
            return response.text().then(function (text) {
              if (!response.ok) {
                if (response.status === 403) {
                  cachedDigest = null;
                }
                throw new Error(
                  "HTTP " + response.status + " " + response.statusText + " from " + target
                );
              }
              if (!parseJson || text === "") {
                return text;
              }
              return JSON.parse(text);
            });
          },
          function (error) {
            if (timer) {
              clearTimeout(timer);
            }
            throw error;
          }
        );
      });
  };

  
  // Export the config and sendRequest function
  export { CONFIG, sendRequest };
  

  // let aa =await fetch("https://tmsobb2.sharepoint.com/_api/web/siteUsers", {
  //   "headers": {
  //     "accept": "application/json",
  //     "accept-language": "en-US,en;q=0.9",
  //     "if-modified-since": "Fri, 18 Apr 2025 15:28:21 GMT",
  //     "priority": "u=1, i",
  //     "sec-ch-ua": "\"Google Chrome\";v=\"135\", \"Not-A.Brand\";v=\"8\", \"Chromium\";v=\"135\"",
  //     "sec-ch-ua-mobile": "?0",
  //     "sec-ch-ua-platform": "\"Windows\"",
  //     "sec-fetch-dest": "empty",
  //     "sec-fetch-mode": "cors",
  //     "sec-fetch-site": "same-origin"
  //   },
  //   "referrer": "https://tmsobb2.sharepoint.com/_layouts/15/people.aspx?MembershipGroupId=50",
  //   "referrerPolicy": "strict-origin-when-cross-origin",
  //   "body": null,
  //   "method": "GET",
  //   "mode": "cors",
  //   "credentials": "include"
  // }).then((g)=>{return g.json()}).then((g)=>{console.log(g);return g})
  // let gid =await aa.value.map((g)=>{return g.Id})

    // let aa =await fetch("https://tmsobb2.sharepoint.com/_api/web/roleDefinitions",
    // let aa =await fetch("https://tmsobb2.sharepoint.com/_api/web/siteGroups",
    