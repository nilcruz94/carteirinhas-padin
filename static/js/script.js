// Exibe e depois esconde a área de flash messages após 3 segundos
setTimeout(function () {
    var flashDiv = document.getElementById('flash-messages');
    if (flashDiv) {
      flashDiv.style.display = 'none';
    }
  }, 3000);
  
  // Função para exibir o overlay de loading
  function showLoading() {
    var existingOverlay = document.getElementById('loading-overlay');
    if (existingOverlay) {
      existingOverlay.remove();
    }
    var loadingOverlay = document.createElement('div');
    loadingOverlay.id = 'loading-overlay';
    loadingOverlay.style.position = 'fixed';
    loadingOverlay.style.top = '0';
    loadingOverlay.style.left = '0';
    loadingOverlay.style.right = '0';
    loadingOverlay.style.bottom = '0';
    loadingOverlay.style.background = 'rgba(0,0,0,0.5)';
    loadingOverlay.style.display = 'flex';
    loadingOverlay.style.alignItems = 'center';
    loadingOverlay.style.justifyContent = 'center';
    loadingOverlay.style.zIndex = '9999';
    loadingOverlay.innerHTML = 
      `<div style="text-align: center; color: white; font-family: Arial, sans-serif;">
        <svg width="3.0cm" height="4.5cm" viewBox="0 0 6.0 9.0" xmlns="http://www.w3.org/2000/svg">
          <rect x="0.3" y="0.3" width="5.4" height="8.4" rx="0.3" ry="0.3" stroke="white" stroke-width="0.1" fill="none" />
          <rect id="badge-fill" x="0.3" y="8.7" width="5.4" height="0" rx="0.3" ry="0.3" fill="white" />
        </svg>
        <p id="loading-text" style="margin-top: 0.2cm;">Gerando carteirinhas...</p>
      </div>`;
    document.body.appendChild(loadingOverlay);
  
    let fillHeight = 0;
    const maxHeight = 8.4;
    function animateBadge() {
      fillHeight += 0.2;
      if (fillHeight > maxHeight) {
        fillHeight = maxHeight;
        clearInterval(interval);
      }
      const badgeFill = document.getElementById('badge-fill');
      badgeFill.setAttribute('y', 8.7 - fillHeight);
      badgeFill.setAttribute('height', fillHeight);
    }
    var interval = setInterval(animateBadge, 100);
    loadingOverlay.dataset.animationId = interval;
  }
  
  // Eventos para mostrar/esconder o formulário de múltiplos uploads
  document.getElementById('show-multi-upload') &&
    document.getElementById('show-multi-upload').addEventListener('click', function () {
      var section = document.getElementById('multi-upload-section');
      if (section.style.display === 'none') {
        section.style.display = 'block';
      } else {
        section.style.display = 'none';
      }
    });
  
  document.getElementById('add-more') &&
    document.getElementById('add-more').addEventListener('click', function () {
      var container = document.getElementById('multi-upload-fields');
      var group = document.createElement('div');
      group.className = 'multi-upload-group';
      group.innerHTML = 
        `<div class="form-group">
           <label>RM do Aluno:</label>
           <input type="text" class="form-control" name="rm[]" placeholder="Digite o RM">
         </div>
         <div class="form-group">
           <input type="file" class="form-control-file" name="foto_file[]" accept="image/*">
         </div>`;
      container.appendChild(group);
    });
  
  // Funções para tratar uploads inline de foto e eventos dos elementos "uploadable"
  document.addEventListener('DOMContentLoaded', function () {
    document.querySelectorAll('.uploadable').forEach(function (element) {
      element.addEventListener('click', function () {
        var rm = element.getAttribute('data-rm');
        var input = document.querySelector('.inline-upload[data-rm="' + rm + '"]');
        if (input) {
          input.click();
        }
      });
    });
  
    document.querySelectorAll('.inline-upload').forEach(function (input) {
      input.addEventListener('change', function () {
        var file = input.files[0];
        if (file) {
          var rm = input.getAttribute('data-rm');
          var formData = new FormData();
          formData.append('rm', rm);
          formData.append('foto_file', file);
          fetch('/upload_inline_foto', {
            method: 'POST',
            body: formData
          })
            .then(response => response.json())
            .then(data => {
              if (data.url) {
                var uploadable = document.querySelector('.uploadable[data-rm="' + rm + '"]');
                if (uploadable.tagName.toLowerCase() === 'img') {
                  uploadable.src = data.url;
                } else {
                  var img = document.createElement('img');
                  img.src = data.url;
                  img.alt = "Foto";
                  img.className = "foto uploadable";
                  img.setAttribute('data-rm', rm);
                  uploadable.parentNode.replaceChild(img, uploadable);
                }
                var msgDiv = document.getElementById('upload-success');
                if (!msgDiv) {
                  msgDiv = document.createElement('div');
                  msgDiv.id = 'upload-success';
                  msgDiv.style.position = 'fixed';
                  msgDiv.style.top = '0.2cm';
                  msgDiv.style.right = '0.2cm';
                  msgDiv.style.backgroundColor = '#d4edda';
                  msgDiv.style.color = '#155724';
                  msgDiv.style.padding = '0.2cm';
                  msgDiv.style.borderRadius = '0.2cm';
                  document.body.appendChild(msgDiv);
                }
                msgDiv.style.display = 'block';
                msgDiv.innerHTML = data.message;
                if (window.flashTimeout) {
                  clearTimeout(window.flashTimeout);
                }
                window.flashTimeout = setTimeout(function () {
                  msgDiv.style.display = 'none';
                }, 3000);
              } else {
                alert("Erro ao fazer upload: " + (data.error || "Erro desconhecido"));
              }
            })
            .catch(error => {
              console.error('Erro:', error);
              alert("Erro no upload da foto.");
            });
        }
      });
    });
  });
  