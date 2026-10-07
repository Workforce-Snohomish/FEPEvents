document.addEventListener("DOMContentLoaded", function () {
  const url =
    "https://workforce-snohomish.github.io/FEPEvents/events.json" // static copy, published by main() in code.gs
  const refreshUrl =
    "https://script.google.com/macros/s/AKfycbzHNhN6wEuTVzf4Fmvs6L23qeZk8vtpLeU9VJ_ZRY1HCx4HubMmOMN7MhrKtIbzKQ7X/exec?refresh=true" // web app, slow, updates the data and the static copy
  const parser = new DOMParser()
  const destination = document.querySelector(".fep-events") // element to drop the fetched data into

  const activeSelectors = new Set() // Store active dataset selectors

  const queryString = Object.fromEntries(
    new URLSearchParams(window.location.search)
  ) // get any queery string parameters
  const refreshRequested = /^true$/i.test(queryString?.refresh) //look for the parameter named refresh 

  // Clear the query string without reloading the page
  window.history.replaceState({}, document.title, window.location.pathname);

  fetch(refreshRequested ? refreshUrl : url)
    .then((response) => {
      if (!response.ok) {
        throw new Error("Network response was not ok " + response.statusText)
      }
      return response.json()
    })
    .then((data) => {
      html = parser.parseFromString(data, "text/html")
      const eventList = html.querySelectorAll(".fep-events > *") // select all elements returned from our API

      while (destination.firstChild) {
        destination.removeChild(destination.firstChild)
      } // remove all, tempoorary elements from the destination object

      destination.append(...eventList)

      const buttons = document.querySelectorAll(".event-type-selector > label")

      buttons.forEach((button) => {
        button.addEventListener("click", (e) => {
          if (e.target !== e.currentTarget) return // Ignore clicks on nested elements

          const selector = e.target.dataset.selector;

          if (activeSelectors.has(selector)) {
            activeSelectors.delete(selector) // Remove if already selected
          } else {
            activeSelectors.add(selector); // Add if not selected
          }
          updateVisible()
        })
      })
    })
    .catch((error) => {
      //TODO: ERROR Handling Required
    })

    function updateVisible() {
        const cards = document.querySelectorAll(".fep-events > .card");
        const classFilter = [...activeSelectors]
            .map((value) => `.${value}`)
            .join("")
        const selectedCards = new Set(
            document.querySelectorAll(`.fep-events > .card${classFilter}`)
        )

        // Use a single loop instead of two
        cards.forEach((card) => {
            card.classList.toggle("event-show", selectedCards.has(card))
        })
    }
})