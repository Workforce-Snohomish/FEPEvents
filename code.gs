/**
 * @fileoverview Google Apps Script backend for the FEP Events widget.
 *
 * Reads upcoming events from the shared FEP Google Calendar, parses the event details out of each
 * description, stores them in the "fepevents" sheet (main, run hourly by a time-driven trigger), and
 * serves the rendered event cards as a JSON-encoded HTML string from the web app (doGet).
 *
 * @version 1.0.0
 * @date 2026-09-18
 */

const fepCalendarID ="c_1883bi678pu1ig4ol917lt3tol4q4@resource.calendar.google.com"

const fields = "items(id, summary, description, start, end)"

const maxEventsToFetch = 50

const properties = [
  'id',
  'summary',
  'description',
  'type',
  'audience',
  'cost',
  'eventURL',
  'start',
  'end',
  'formattedTime',
  'internal',
  'cssSelector'
]

const propertyIndex = properties.reduce((acc, property, i) => {acc[property] = i;return acc}, {})
const sheetName = "fepevents"
const timeZone = Session.getScriptTimeZone() // read the timezone under which this script is running

function toCamelCase(value) {
    // prototype to convert a string to camelCase
    return value
          .trim()
          .toLowerCase()
          .replace(/[^a-zA-Z0-9]+(.)/g, (m, chr) => chr.toUpperCase())
}

function toKebabCase(value) {
  return 'event-' + value
        .trim()                // Remove leading/trailing spaces
        .toLowerCase()         // Convert to lowercase
        .replace(/\s+/g, "-")  // Replace spaces with hyphens
        .replace(/[^a-z0-9-]/g, "") // Remove non-alphanumeric characters (except hyphens)
}

function toTitleCase(value) {
    return value
        .trim()
        .toLowerCase()
        .replace(/-/g," ") // remove hyphens
        .replace(/\b\w/g, char => char.toUpperCase())
}

function sameDay(dtstart, dtend) {
  return Utilities.formatDate(dtstart, timeZone, "yyyyMMdd") == Utilities.formatDate(dtend, timeZone, "yyyyMMdd")
}

function convertEventsToArray(events) {
  const eventArray = events.map(obj => {
      const row = [] // temporary array used to build each rows data
      properties.forEach(function(property) {
        if (obj[property] === undefined) {
            row.push("")
        } else {
            row.push(obj[property])
        }
      })
      return row
  })
  eventArray.unshift(properties) // add the list of properties themselves as the first row in the array
  return eventArray
}

function saveData(events) {
  const ss = SpreadsheetApp.getActive()
  const outputSheet = ss.getSheetByName(sheetName) ?? ss.insertSheet(sheetName, 1)

  outputSheet.getDataRange().clear()
  outputSheet.getRange(1, 1, events.length, events[0].length).setValues(events)
}

function check() {
  const today = new Date() // date from which to request events
  console.log(Utilities.formatDate(today, timeZone, "yyyy-MM-dd\'T\'00:00:00Z"))
  const testDate = new Date("2025-02-27T20:00:00-08:00")
  console.log(testDate)
  const description = "\u003cb\u003eAudience: \u003c/b\u003eGeneral Public\u003cbr\u003e\u003cbr\u003e\u003cb\u003eType: \u003c/b\u003eVirtual\u003cb\u003e\u003cbr\u003e\u003c/b\u003e\u003cbr\u003e\u003cb\u003eCost: \u003c/b\u003eFree\u003cbr\u003e\u003cbr\u003e\u003cb\u003eEvent: \u003c/b\u003eThe Financial Literacy and Education Commission was established under the Fair and Accurate Credit Transactions Act of 2003. The Commission was tasked to develop a national financial education web site (MyMoney.gov) and a national strategy on financial education. It is chaired by the Secretary of the Treasury and the vice chair is the Director of the Bureau of Consumer Financial Protection. The Commission is coordinated by the Department of the Treasury's Office of Consumer Policy.\u003cbr\u003e\u003cbr\u003eThe Commission’s vision is of sustained financial well-being for all individuals and families in the U.S. In furtherance of this vision, the Commission sets strategic direction for policy, education, practice, research, and coordination so that all Americans make informed financial decisions\u003cbr\u003e\u003cbr\u003eRegistration is not required.\u003cbr\u003e\u003cbr\u003e\u003cb\u003eEvent URL: \u003c/b\u003ehttps://links-2.govdelivery.com/CL0/https:%2F%2Fusdotyorktel.rev.vbrick.com%2F%23%2Fevents%2F1a1024d1-8af7-4b24-bf5a-43590313b496/1/0101019f66a2a602-79fc3420-6a9b-4272-a52b-f856ba7586c5-000000/Yh5xVaI2Ry_1TtEfyMF6VySFB09fub8EQgCABi7mgss=452"
  console.log(decodeURI(description))
}

function main() {
  const fepEvents = getEvents()
  const parsedEvents = parseEvents(fepEvents)
  saveData(convertEventsToArray(parsedEvents))
  try {
    publishToGitHub(JSON.stringify(renderEvents()))
  } catch (e) {
    console.log(`Publishing to GitHub failed: ${e}`) // the sheet is still up to date, so don't fail the run
  }
}

/**
 * Commits the rendered events to a file in a GitHub repo, from where GitHub Pages serves it to client.js.
 * Configured through Script Properties: GITHUB_TOKEN, GITHUB_REPO ("owner/name") and optionally
 * GITHUB_BRANCH (default "gh-pages") and GITHUB_PATH (default "events.json").
 * Nothing is committed when the content is unchanged.
 */
function publishToGitHub(content) {
  const scriptProperties = PropertiesService.getScriptProperties()
  const token = scriptProperties.getProperty('GITHUB_TOKEN')
  const repo = scriptProperties.getProperty('GITHUB_REPO')
  const branch = scriptProperties.getProperty('GITHUB_BRANCH') || 'gh-pages'
  const path = scriptProperties.getProperty('GITHUB_PATH') || 'events.json'

  if (!token || !repo) {
    console.log('GITHUB_TOKEN or GITHUB_REPO not set, skipping publish')
    return
  }

  const apiUrl = `https://api.github.com/repos/${repo}/contents/${path}`
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  }

  // fetch the current file, both to compare it and to get the sha GitHub requires for an update
  const current = UrlFetchApp.fetch(`${apiUrl}?ref=${encodeURIComponent(branch)}`, {headers: headers, muteHttpExceptions: true})
  let sha
  if (current.getResponseCode() == 200) {
    const file = JSON.parse(current.getContentText())
    sha = file.sha
    const currentContent = Utilities.newBlob(Utilities.base64Decode(file.content.replace(/\s/g, ""))).getDataAsString()
    if (currentContent === content) {
      console.log('Published events are already up to date')
      return
    }
  } else if (current.getResponseCode() != 404) {
    throw new Error(`GitHub read failed (${current.getResponseCode()}): ${current.getContentText()}`)
  }

  const payload = {
    message: `Update events ${Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd HH:mm")}`,
    content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
    branch: branch
  }
  if (sha) payload.sha = sha

  const response = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    headers: headers,
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  })
  if (response.getResponseCode() >= 300) {
    throw new Error(`GitHub write failed (${response.getResponseCode()}): ${response.getContentText()}`)
  }
  console.log(`Published events to ${repo}/${path} on ${branch}`)
}

/**
 * Run by hand from the editor to find out why publishing fails: logs whether the token is accepted,
 * whether it can see and write to the repo, and whether the branch exists.
 */
function checkGitHubSetup() {
  const scriptProperties = PropertiesService.getScriptProperties()
  const token = scriptProperties.getProperty('GITHUB_TOKEN')
  const repo = scriptProperties.getProperty('GITHUB_REPO')
  const branch = scriptProperties.getProperty('GITHUB_BRANCH') || 'gh-pages'
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  }
  const get = url => UrlFetchApp.fetch(url, {headers: headers, muteHttpExceptions: true})

  console.log(`GITHUB_REPO is "${repo}", branch is "${branch}", token ${token ? `is set (${token.length} characters)` : 'is NOT set'}`)

  const repoResponse = get(`https://api.github.com/repos/${repo}`)
  console.log(`Repo lookup: ${repoResponse.getResponseCode()}`)
  if (repoResponse.getResponseCode() == 401) {
    console.log('The token itself was rejected: it is mistyped, expired or revoked')
    return
  }
  if (repoResponse.getResponseCode() != 200) {
    console.log('The token cannot see this repo: check the GITHUB_REPO spelling (owner/name), that the token was created for this owner and repo, and that the organization has approved it')
    return
  }
  const repoDetails = JSON.parse(repoResponse.getContentText())
  console.log(`Repo found: ${repoDetails.full_name}, ${repoDetails.private ? 'PRIVATE' : 'public'}, default branch ${repoDetails.default_branch}`)
  console.log(`Token can write contents: ${repoDetails.permissions ? repoDetails.permissions.push : 'unknown'}`)

  const branchResponse = get(`https://api.github.com/repos/${repo}/branches/${encodeURIComponent(branch)}`)
  console.log(`Branch "${branch}" lookup: ${branchResponse.getResponseCode()}${branchResponse.getResponseCode() == 200 ? ' (exists)' : ' (does not exist yet, it needs to be pushed or created)'}`)
}

function getEvents() {
  const today = new Date() // date from which to request events

  const queryParameters = {
    timeMin: Utilities.formatDate(today, timeZone, "yyyy-MM-dd\'T\'00:00:00Z"),
    fields: fields,
    singleEvents: true,
    maxResults: maxEventsToFetch,
    orderBy: "startTime"
  }
  const fepEvents = Calendar.Events.list(
    fepCalendarID,
    queryParameters
  )
  return fepEvents.items || []
}

function parseEvents(events) {
  const fepEvents = [] // create an array to store the details of each event
  //const regExpRegister = new RegExp(/(?:Event URL\s*:\s*)(?:<[^>]+>)*(.*?)href\s*=\s*["']([^"']+)["']/i) // used to extract the Event URL if present
  const regExpRegister = new RegExp(/(?:(?:<[^>]*>)|\s)*Event\s*URL(?:(?:<[^>]*>)|\s)*:(?:(?:<[^>]*>)|\s)*(?:<a\s+[^>]*href\s*=\s*["']([^"']+)["']|(\bhttps?:\/\/[^\s<]+))/i)
  /*
  Pattern breakdown:
  - (?:(?:<[^>]*>)|\s)*     // Optional tags or whitespace before "Event URL"
  - Event\s*URL            // Matches "Event URL" with optional whitespace between words
  - (?:(?:<[^>]*>)|\s)*    // Optional tags/whitespace between "Event URL" and colon
  - :                      // The literal colon character
  - (?:(?:<[^>]*>)|\s)*    // Optional tags/whitespace after the colon
  - <a\s+[^>]*             // Opening <a tag with attributes
  - href\s*=\s*            // href attribute with optional whitespace
  - ["']([^"']+)["']       // Capture group 1: URL within quotes (single or double)

  Handles cases like:
  - <b>Event URL: </b><a href="...">
  - <b>Event URL</b><b>:</b><a href="...">
  - Event URL: <a href="...">
  - <b>Event</b> <b>URL</b><b>:</b> <a href="...">

  Usage:
  const match = eventUrlRegex.exec(rawCalendarData);
  const eventUrl = match ? match[1] : null;
  */
  
  const regExpInternal = new RegExp(/^\[Int\][^\w]*/i) // tag marking internal events
  const regExpAudience = new RegExp(/(?:Audience\s*:\s*)(?:<[^>]+>)*(.*?)(?=<[^>]+>)/i)
  const regExpType = new RegExp(/(?:Type\s*:\s*)(?:<[^>]+>)*(.*?)(?=<[^>]+>)/i)
  const regExpCost = new RegExp(/(?:Cost\s*:\s*)(?:<[^>]+>)*(.*?)(?=<[^>]+>)/i)
  //const regExpDescription = new RegExp(/(?:Event[^:]*:(?:\s*<[^>]+>)?\W*)(?:<[^>]+>)*(.*?)(?=<br>)/i)

  //const regExpDescription = new RegExp(/(?:Event[^:]*:(?:\s*(?:<[^>]*>)*)?\W*)(?:<[^>]+>)*(.*?)(?=<b>[\w\s][^:]*)/i)
  const regExpDescription = new RegExp(/(?:Event[^:]*:\s*(?:<[^>]*>\s*)*)(.*?)(?=<b>[\w\s]+(?:<[^>]*>)?:|$)/i)

  let eventURL, internal, type, audience, description, cost

  events.forEach((event,i)  => {
    try {
    //console.log(`Event number ${i} Raw description: ${event.description}`)
    //event.description = decodeURI(event.description) // decode the uri encoded data in the description
    eventURL = event.description.match(regExpRegister)
    //console.log(`Event number ${i}, URL: ${eventURL??"N/A"}`)
    internal = event.summary.match(regExpInternal)
    description = event.description.match(regExpDescription)
    //console.log(`Event number ${i} description: ${description}`)
    audience = event.description.match(regExpAudience)
    type = event.description.match(regExpType)
    cost = event.description.match(regExpCost)
    } catch (e) {
      console.log(`Event number ${i}, error ${e}`)
    }
    

    const currentEvent = {
      id: event.id,
      summary: event.summary.replace(regExpInternal,''),
      start: new Date(event.start.dateTime),
      end: new Date(event.end.dateTime),
      eventURL: eventURL ? (eventURL[1] || eventURL[2]) : null,
      internal: internal ? true : false,
      description: description ? description[1].replace(/<\/*i>|<\/*b>|<\/*strong>|<\/*span>/gi,"") : null,
      type: type ? type[1] : null,
      cost: cost ? cost[1] : null,
      audience: audience ? audience[1] : null
    }

    currentEvent.cssSelector = `${currentEvent.audience ? toKebabCase(currentEvent.audience) : ""} ${currentEvent.cost ? toKebabCase(currentEvent.cost) : ""} ${currentEvent.type ? toKebabCase(currentEvent.type) : ""}`

    if (sameDay(currentEvent.start, currentEvent.end)) {
      currentEvent.formattedTime = Utilities.formatDate(currentEvent.start, timeZone, "EEEE MMMM d, yyyy h:mm a") + " - " + Utilities.formatDate(currentEvent.end, timeZone, "h:mm a")
    } else {
      currentEvent.formattedTime = Utilities.formatDate(currentEvent.start, timeZone, "EEEE MMMM d, yyyy h:mm a") + " - " + Utilities.formatDate(currentEvent.end, timeZone, "EEEE MMMM d, yyyy h:mm a")
    }
    fepEvents.push(currentEvent)
  })
  return fepEvents
}

function doGet(e) {
  if (e!= null && e?.parameter?.refresh) {
    main() // update the data in the spreadsheet if the parameter refresh=true is present in the querystring
  }

  return ContentService.createTextOutput(JSON.stringify(renderEvents()))
    .setMimeType(ContentService.MimeType.JSON)
}

/**
 * Renders the events currently stored in the sheet as the HTML for the widget.
 */
function renderEvents() {
  const ss = SpreadsheetApp.getActive()
  const sd = ss.getSheetByName(sheetName)
  const data = sd.getDataRange().getDisplayValues()
  const template = HtmlService.createTemplateFromFile('HtmlTemplate')

  data.shift() // remove the header row from the data before passing it to the template

  const uniqueFilters = new Set(
    data
      .flatMap(row => row[propertyIndex.cssSelector]
        .split(" ")
        .filter(value => value !== "")
      )
  )  // create a unique list of filters which will then be used to create the necessary filter buttons

  template.data = data
  template.propertyIndex = propertyIndex

  template.uniqueFilters = [...uniqueFilters]
    .sort((a,b) => {return a.localeCompare(b)})
    .map(value => ({
      cssSelector: value,
      buttonText: toTitleCase(value.replace(/^event-/i,""))
  }))

  return template.evaluate().getContent()
}
