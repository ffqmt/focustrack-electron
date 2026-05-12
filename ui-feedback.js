(function () {
  function ensureDialogRoot() {
    let root = document.getElementById('ftDialogRoot');

    if (!root) {
      root = document.createElement('div');
      root.id = 'ftDialogRoot';
      document.body.appendChild(root);
    }

    return root;
  }

  function getDialogMeta(type) {
    switch (type) {
      case 'success':
        return {
          icon: '✓',
          title: 'Sucesso',
          className: 'ft-dialog--success'
        };

      case 'error':
        return {
          icon: '!',
          title: 'Atenção',
          className: 'ft-dialog--error'
        };

      case 'warning':
        return {
          icon: '!',
          title: 'Atenção',
          className: 'ft-dialog--warning'
        };

      case 'info':
      default:
        return {
          icon: 'i',
          title: 'Informação',
          className: 'ft-dialog--info'
        };
    }
  }

  function escapeHtml(value) {
    return String(value || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  function closeDialog(dialogEl, resolve, value) {
    if (!dialogEl) {
      resolve(value);
      return;
    }

    dialogEl.classList.add('ft-dialog-shell--closing');

    window.setTimeout(() => {
      dialogEl.remove();
      resolve(value);
    }, 140);
  }

  function showDialog(options = {}) {
    const root = ensureDialogRoot();

    const type = options.type || 'info';
    const meta = getDialogMeta(type);

    const title = options.title || meta.title;
    const message = options.message || '';
    const confirmText = options.confirmText || 'OK';
    const cancelText = options.cancelText || 'Cancelar';
    const showCancel = Boolean(options.showCancel);

    return new Promise((resolve) => {
      const shell = document.createElement('div');

      shell.className = `ft-dialog-shell ${meta.className}`;
      shell.innerHTML = `
        <div class="ft-dialog-backdrop" data-ft-dialog-close="backdrop"></div>

        <div class="ft-dialog-card" role="dialog" aria-modal="true">
          <div class="ft-dialog-header">
            <div class="ft-dialog-icon">
              ${escapeHtml(meta.icon)}
            </div>

            <div class="ft-dialog-title-area">
              <h2 class="ft-dialog-title">${escapeHtml(title)}</h2>
              ${
                message
                  ? `<p class="ft-dialog-message">${escapeHtml(message)}</p>`
                  : ''
              }
            </div>
          </div>

          <div class="ft-dialog-actions">
            ${
              showCancel
                ? `
                  <button
                    class="ft-dialog-btn ft-dialog-btn--neutral"
                    type="button"
                    data-ft-dialog-action="cancel"
                  >
                    ${escapeHtml(cancelText)}
                  </button>
                `
                : ''
            }

            <button
              class="ft-dialog-btn ft-dialog-btn--primary"
              type="button"
              data-ft-dialog-action="confirm"
            >
              ${escapeHtml(confirmText)}
            </button>
          </div>
        </div>
      `;

      root.appendChild(shell);

      requestAnimationFrame(() => {
        shell.classList.add('ft-dialog-shell--visible');
      });

      const confirmBtn = shell.querySelector('[data-ft-dialog-action="confirm"]');
      const cancelBtn = shell.querySelector('[data-ft-dialog-action="cancel"]');

      confirmBtn?.focus?.();

      confirmBtn?.addEventListener('click', () => {
        closeDialog(shell, resolve, true);
      });

      cancelBtn?.addEventListener('click', () => {
        closeDialog(shell, resolve, false);
      });

      shell.addEventListener('click', (event) => {
        const target = event.target;

        if (target?.dataset?.ftDialogClose === 'backdrop' && showCancel) {
          closeDialog(shell, resolve, false);
        }
      });

      function onKeyDown(event) {
        if (!document.body.contains(shell)) {
          document.removeEventListener('keydown', onKeyDown);
          return;
        }

        if (event.key === 'Enter') {
          event.preventDefault();
          closeDialog(shell, resolve, true);
        }

        if (event.key === 'Escape') {
          event.preventDefault();
          closeDialog(shell, resolve, showCancel ? false : true);
        }
      }

      document.addEventListener('keydown', onKeyDown);
    });
  }

  function showToast(options = {}) {
    const root = ensureDialogRoot();

    const type = options.type || 'info';
    const meta = getDialogMeta(type);
    const message = options.message || '';
    const duration = Number(options.duration || 2800);

    const toast = document.createElement('div');
    toast.className = `ft-toast ${meta.className}`;
    toast.innerHTML = `
      <span class="ft-toast-icon">${escapeHtml(meta.icon)}</span>
      <span class="ft-toast-message">${escapeHtml(message)}</span>
    `;

    root.appendChild(toast);

    requestAnimationFrame(() => {
      toast.classList.add('ft-toast--visible');
    });

    window.setTimeout(() => {
      toast.classList.remove('ft-toast--visible');
      toast.classList.add('ft-toast--closing');

      window.setTimeout(() => {
        toast.remove();
      }, 160);
    }, duration);
  }

  window.FocusTrackUI = {
    alert(options = {}) {
      return showDialog({
        ...options,
        showCancel: false
      });
    },

    confirm(options = {}) {
      return showDialog({
        ...options,
        showCancel: true
      });
    },

    toast(options = {}) {
      return showToast(options);
    }
  };
})();
