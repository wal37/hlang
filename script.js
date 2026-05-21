document.querySelectorAll('form[action="#"]').forEach((form) => {
  form.addEventListener("submit", (event) => {
    event.preventDefault();
    const button = form.querySelector('button[type="submit"]');
    if (!button) return;

    const original = button.textContent;
    button.textContent = "Received";
    button.disabled = true;

    window.setTimeout(() => {
      button.textContent = original;
      button.disabled = false;
      form.reset();
    }, 1800);
  });
});

document.querySelectorAll('.nav-toggle').forEach((button) => {
  const targetId = button.getAttribute("aria-controls");
  const nav = targetId ? document.getElementById(targetId) : null;
  if (!nav) return;

  button.addEventListener("click", () => {
    const isOpen = button.getAttribute("aria-expanded") === "true";
    button.setAttribute("aria-expanded", String(!isOpen));
    nav.classList.toggle("is-open", !isOpen);
  });

  nav.querySelectorAll("a").forEach((link) => {
    link.addEventListener("click", () => {
      button.setAttribute("aria-expanded", "false");
      nav.classList.remove("is-open");
    });
  });
});
