// 확장 행 토글
function toggleExpand(btn) {
  const row = btn.closest ? btn.closest('tr') : btn;
  const next = row.nextElementSibling;
  if (!next || !next.classList.contains('expand-row')) return;
  next.classList.toggle('open');
  if (btn.textContent !== undefined)
    btn.textContent = next.classList.contains('open') ? '▴' : '▾';
}

// select auto-submit
document.addEventListener('DOMContentLoaded', function () {
  document.querySelectorAll('.auto-submit select').forEach(function (el) {
    el.addEventListener('change', function () { el.closest('form').submit(); });
  });
});
