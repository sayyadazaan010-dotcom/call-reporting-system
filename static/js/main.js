// ============================================================
// Futuristic Network Background (nodes + connecting lines)
// ============================================================
function initNetworkBackground() {
  const container = document.querySelector('.bg-fx');
  if (!container) return;

  // Clear old particles if any
  container.innerHTML = '';

  const canvas = document.createElement('canvas');
  canvas.id = 'netCanvas';
  canvas.style.cssText = 'position:absolute;inset:0;width:100%;height:100%;';
  container.appendChild(canvas);

  const ctx = canvas.getContext('2d');
  let w, h, nodes = [], animId;
  const NODE_COUNT = window.innerWidth < 700 ? 45 : 80;
  const CONNECT_DIST = 140;
  const MOUSE_DIST = 160;

  let mouse = { x: null, y: null };

  function resize() {
    w = canvas.width = window.innerWidth;
    h = canvas.height = window.innerHeight;
  }

  function createNodes() {
    nodes = [];
    for (let i = 0; i < NODE_COUNT; i++) {
      nodes.push({
        x: Math.random() * w,
        y: Math.random() * h,
        vx: (Math.random() - 0.5) * 0.55,
        vy: (Math.random() - 0.5) * 0.55,
        r: Math.random() * 1.8 + 0.8,
        pulse: Math.random() * Math.PI * 2,
      });
    }
  }

  function draw() {
    ctx.clearRect(0, 0, w, h);

    // soft radial glow behind center
    const grad = ctx.createRadialGradient(w * 0.5, h * 0.35, 0, w * 0.5, h * 0.35, Math.max(w, h) * 0.55);
    grad.addColorStop(0, 'rgba(59,130,246,0.06)');
    grad.addColorStop(0.5, 'rgba(139,92,246,0.03)');
    grad.addColorStop(1, 'transparent');
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, w, h);

    // update + draw nodes
    for (let i = 0; i < nodes.length; i++) {
      const n = nodes[i];
      n.x += n.vx;
      n.y += n.vy;
      n.pulse += 0.02;

      // bounce
      if (n.x < 0 || n.x > w) n.vx *= -1;
      if (n.y < 0 || n.y > h) n.vy *= -1;

      // mild mouse attraction
      if (mouse.x !== null) {
        const dx = mouse.x - n.x;
        const dy = mouse.y - n.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < MOUSE_DIST) {
          n.vx += dx * 0.00015;
          n.vy += dy * 0.00015;
        }
      }

      // limit speed
      const sp = Math.sqrt(n.vx * n.vx + n.vy * n.vy);
      if (sp > 1.2) {
        n.vx = (n.vx / sp) * 1.2;
        n.vy = (n.vy / sp) * 1.2;
      }

      const glow = 0.45 + 0.35 * Math.sin(n.pulse);
      ctx.beginPath();
      ctx.arc(n.x, n.y, n.r, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(34, 211, 238, ${glow})`;
      ctx.fill();

      // outer soft glow
      ctx.beginPath();
      ctx.arc(n.x, n.y, n.r * 3.5, 0, Math.PI * 2);
      ctx.fillStyle = `rgba(59, 130, 246, ${glow * 0.12})`;
      ctx.fill();
    }

    // connections
    for (let i = 0; i < nodes.length; i++) {
      for (let j = i + 1; j < nodes.length; j++) {
        const a = nodes[i];
        const b = nodes[j];
        const dx = a.x - b.x;
        const dy = a.y - b.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < CONNECT_DIST) {
          const alpha = (1 - dist / CONNECT_DIST) * 0.35;
          ctx.beginPath();
          ctx.moveTo(a.x, a.y);
          ctx.lineTo(b.x, b.y);
          ctx.strokeStyle = `rgba(99, 140, 255, ${alpha})`;
          ctx.lineWidth = 0.8;
          ctx.stroke();
        }
      }
    }

    // mouse connections
    if (mouse.x !== null) {
      for (let i = 0; i < nodes.length; i++) {
        const n = nodes[i];
        const dx = mouse.x - n.x;
        const dy = mouse.y - n.y;
        const dist = Math.sqrt(dx * dx + dy * dy);
        if (dist < MOUSE_DIST) {
          const alpha = (1 - dist / MOUSE_DIST) * 0.5;
          ctx.beginPath();
          ctx.moveTo(mouse.x, mouse.y);
          ctx.lineTo(n.x, n.y);
          ctx.strokeStyle = `rgba(34, 211, 238, ${alpha})`;
          ctx.lineWidth = 1;
          ctx.stroke();
        }
      }
    }

    animId = requestAnimationFrame(draw);
  }

  function onMove(e) {
    mouse.x = e.clientX;
    mouse.y = e.clientY;
  }
  function onLeave() {
    mouse.x = null;
    mouse.y = null;
  }
  function onTouch(e) {
    if (e.touches && e.touches[0]) {
      mouse.x = e.touches[0].clientX;
      mouse.y = e.touches[0].clientY;
    }
  }

  resize();
  createNodes();
  draw();

  window.addEventListener('resize', () => {
    resize();
    createNodes();
  });
  window.addEventListener('mousemove', onMove);
  window.addEventListener('mouseleave', onLeave);
  window.addEventListener('touchmove', onTouch, { passive: true });
  window.addEventListener('touchend', onLeave);
}

// Toggle-group buttons (Remote/Visit, Pending/Completed) driving a hidden input
function initToggleGroups() {
  document.querySelectorAll('[data-toggle-group]').forEach((group) => {
    const hiddenInput = document.querySelector(group.dataset.target);
    group.querySelectorAll('.toggle-option').forEach((opt) => {
      opt.addEventListener('click', () => {
        group.querySelectorAll('.toggle-option').forEach((o) => o.classList.remove('active'));
        opt.classList.add('active');
        if (hiddenInput) hiddenInput.value = opt.dataset.value;
      });
    });
  });
}

// Simple modal open/close helpers
function openModal(id) {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.add('show');
  // Click on dark backdrop closes modal
  el.onclick = (e) => {
    if (e.target === el) closeModal(id);
  };
}
function closeModal(id) {
  document.getElementById(id)?.classList.remove('show');
}

// Auto-dismiss flash messages
function initFlashAutoHide() {
  document.querySelectorAll('.flash').forEach((el) => {
    setTimeout(() => {
      el.style.transition = 'opacity 0.5s ease';
      el.style.opacity = '0';
      setTimeout(() => el.remove(), 500);
    }, 4500);
  });
}

document.addEventListener('DOMContentLoaded', () => {
  initNetworkBackground();
  initToggleGroups();
  initFlashAutoHide();
});
