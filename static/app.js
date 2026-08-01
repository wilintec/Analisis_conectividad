(() => {
  "use strict";

  const slideData = JSON.parse(document.getElementById("slide-data").textContent);
  const cards = [...document.querySelectorAll(".slide-card")];
  const chapters = [...document.querySelectorAll(".chapter")];
  const search = document.getElementById("search");
  const noResults = document.getElementById("no-results");
  const dialog = document.getElementById("slide-dialog");
  const dialogImage = document.getElementById("dialog-image");
  const dialogNumber = document.getElementById("dialog-number");
  const dialogTitle = document.getElementById("dialog-title");
  let currentSlide = 1;

  function normalize(value) {
    return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  }

  function filterSlides() {
    const query = normalize(search.value);
    let visibleCount = 0;
    cards.forEach((card) => {
      const haystack = normalize(card.dataset.search || "");
      const visible = !query || haystack.includes(query);
      card.hidden = !visible;
      if (visible) visibleCount += 1;
    });
    chapters.forEach((chapter) => {
      chapter.hidden = !chapter.querySelector(".slide-card:not([hidden])");
    });
    noResults.hidden = visibleCount !== 0;
  }

  function openSlide(number) {
    const slide = slideData.find((item) => item.number === number);
    if (!slide) return;
    currentSlide = number;
    dialogImage.src = slide.image;
    dialogImage.alt = `Diapositiva ${number}: ${slide.title}`;
    dialogNumber.textContent = String(number).padStart(2, "0");
    dialogTitle.textContent = slide.title;
    if (!dialog.open) dialog.showModal();
  }

  function changeSlide(delta) {
    const next = Math.min(slideData.length, Math.max(1, currentSlide + delta));
    openSlide(next);
  }

  search.addEventListener("input", filterSlides);
  document.querySelectorAll(".open-slide").forEach((button) => {
    button.addEventListener("click", () => openSlide(Number(button.dataset.slide)));
  });
  document.querySelector(".dialog-close").addEventListener("click", () => dialog.close());
  document.querySelector(".dialog-nav.previous").addEventListener("click", () => changeSlide(-1));
  document.querySelector(".dialog-nav.next").addEventListener("click", () => changeSlide(1));
  dialog.addEventListener("click", (event) => {
    if (event.target === dialog) dialog.close();
  });
  document.addEventListener("keydown", (event) => {
    if (!dialog.open) return;
    if (event.key === "ArrowLeft") changeSlide(-1);
    if (event.key === "ArrowRight") changeSlide(1);
  });

  const menuButton = document.querySelector(".menu-button");
  const chapterNav = document.getElementById("chapter-nav");
  menuButton.addEventListener("click", () => {
    const open = chapterNav.classList.toggle("open");
    menuButton.setAttribute("aria-expanded", String(open));
  });
  chapterNav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => {
      chapterNav.classList.remove("open");
      menuButton.setAttribute("aria-expanded", "false");
    });
  });
})();
