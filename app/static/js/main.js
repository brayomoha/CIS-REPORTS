/* CIS School System — Main JavaScript */

// ============================================================
// MARK INPUT — auto colour-code based on value
// ============================================================
function colourMarkInput(input) {
  const val = parseFloat(input.value);
  input.classList.remove('ee','me','ae','be');
  if (isNaN(val)) return;
  if (val >= 26) input.classList.add('ee');
  else if (val >= 18) input.classList.add('me');
  else if (val >= 11) input.classList.add('ae');
  else input.classList.add('be');
}

document.querySelectorAll('.mark-input').forEach(inp => {
  inp.addEventListener('input', () => colourMarkInput(inp));
  colourMarkInput(inp);  // run on load for pre-filled values
});

// ============================================================
// RECEPTION — rating picker buttons
// ============================================================
document.querySelectorAll('.rating-picker').forEach(picker => {
  const buttons = picker.querySelectorAll('.rating-btn');
  const hiddenInput = picker.querySelector('input[type=hidden]');

  buttons.forEach(btn => {
    btn.addEventListener('click', () => {
      const val = btn.dataset.val;
      // Deselect all
      buttons.forEach(b => b.className = 'rating-btn');
      // Select clicked
      btn.classList.add(`selected-${val}`);
      if (hiddenInput) hiddenInput.value = val;
    });
  });

  // Highlight pre-saved value on load
  if (hiddenInput && hiddenInput.value) {
    const pre = picker.querySelector(`[data-val="${hiddenInput.value}"]`);
    if (pre) pre.classList.add(`selected-${hiddenInput.value}`);
  }
});

// ============================================================
// COMMENT BANK PICKER — click to fill textarea
// ============================================================
document.querySelectorAll('.comment-option').forEach(opt => {
  opt.addEventListener('click', () => {
    const target = document.getElementById(opt.dataset.target);
    if (target) {
      target.value = opt.dataset.text;
      target.focus();
      // Visual feedback
      opt.style.background = 'var(--gold-pale)';
      setTimeout(() => opt.style.background = '', 600);
    }
  });
});

// ============================================================
// FLASH MESSAGE — auto dismiss after 5s
// ============================================================
document.querySelectorAll('.flash').forEach(flash => {
  setTimeout(() => {
    flash.style.opacity = '0';
    flash.style.transition = 'opacity 0.4s';
    setTimeout(() => flash.remove(), 400);
  }, 5000);
});

// ============================================================
// MARK SHEET — Tab key moves to next input in row
// ============================================================
const markInputs = Array.from(document.querySelectorAll('.mark-input'));
markInputs.forEach((inp, idx) => {
  inp.addEventListener('keydown', e => {
    if (e.key === 'Tab' && !e.shiftKey) {
      e.preventDefault();
      const next = markInputs[idx + 1];
      if (next) next.focus();
    }
  });
});

// ============================================================
// CONFIRM DELETE / DANGEROUS ACTIONS
// ============================================================
document.querySelectorAll('[data-confirm]').forEach(el => {
  el.addEventListener('click', e => {
    if (!confirm(el.dataset.confirm)) e.preventDefault();
  });
});

// ============================================================
// SIDEBAR — close when clicking outside on mobile
// ============================================================
document.addEventListener('click', e => {
  const sidebar = document.getElementById('sidebar');
  const toggle  = document.querySelector('.menu-toggle');
  if (sidebar && sidebar.classList.contains('open')
      && !sidebar.contains(e.target)
      && e.target !== toggle) {
    sidebar.classList.remove('open');
  }
});

// ============================================================
// PROGRESS BARS — animate on load
// ============================================================
document.querySelectorAll('.progress-bar').forEach(bar => {
  const target = bar.dataset.width || '0';
  bar.style.width = '0%';
  requestAnimationFrame(() => {
    setTimeout(() => { bar.style.width = target + '%'; }, 100);
  });
  if (parseInt(target) === 100) bar.classList.add('complete');
});

// ============================================================
// IMPORT MODAL (students)
// ============================================================
const importBtn = document.getElementById('importBtn');
const importModal = document.getElementById('importModal');
if (importBtn && importModal) {
  importBtn.addEventListener('click', () => importModal.classList.toggle('hidden'));
  importModal.addEventListener('click', e => {
    if (e.target === importModal) importModal.classList.add('hidden');
  });
}
