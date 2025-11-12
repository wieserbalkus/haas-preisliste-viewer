(function(){
  'use strict';

  const COMPLETION_THRESHOLD = 0.99;
  const METADATA_RETRY_LIMIT = 4;
  const METADATA_RETRY_DELAY = 1200;
  const STATE = {
    IDLE: 'idle',
    DOWNLOADING: 'downloading',
    READY: 'ready',
    PLAYING: 'playing',
    PAUSED: 'paused',
    ERROR: 'error'
  };

  function formatTime(seconds){
    if(!Number.isFinite(seconds) || seconds < 0){
      return '0:00';
    }
    const totalSeconds = Math.floor(seconds);
    const m = Math.floor(totalSeconds / 60);
    const s = totalSeconds % 60;
    return `${m}:${String(s).padStart(2,'0')}`;
  }

  function formatBytes(bytes){
    if(!Number.isFinite(bytes) || bytes <= 0){
      return '';
    }
    const units = ['B','KB','MB','GB'];
    let idx = 0;
    let value = bytes;
    while(value >= 1024 && idx < units.length - 1){
      value /= 1024;
      idx += 1;
    }
    return `${value.toFixed(value >= 10 || idx === 0 ? 0 : 1)} ${units[idx]}`;
  }

  class AudioCardController{
    constructor(card){
      this.card = card;
      this.source = card.dataset.audioSrc || '';
      if(!this.source){
        return;
      }

      this.ensureStructure();

      this.actionBtn = card.querySelector('[data-role="audio-action"]');
      this.toggleBtn = card.querySelector('[data-role="audio-toggle"]');
      this.playerWrap = card.querySelector('[data-role="audio-player"]');
      this.seekInput = card.querySelector('[data-role="audio-seek"]');
      this.currentLabel = card.querySelector('[data-role="audio-current"]');
      this.durationLabels = card.querySelectorAll('[data-role="audio-duration"]');
      this.sizeLabel = card.querySelector('[data-role="audio-size"]');
      this.statusLabel = card.querySelector('[data-role="audio-status"]');
      this.completionLabel = card.querySelector('[data-role="audio-complete"]');

      this.audio = new Audio();
      this.audio.preload = 'auto';
      this.audio.crossOrigin = card.dataset.audioCrossorigin || 'anonymous';

      this.state = STATE.IDLE;
      this.isComplete = false;
      this.duration = 0;
      this.metadataReady = false;
      this.metadataAttempts = 0;
      this.metadataTimer = null;
      this.seekFromInput = false;
      this.playerVisible = false;
      this.downloadPromise = null;
      this.objectUrl = null;
      this.cleanupHandler = () => this.cleanup();

      this.card.dataset.audioComplete = 'false';

      window.addEventListener('beforeunload', this.cleanupHandler);

      this.bindEvents();
      this.updateActionButton();
      this.updateStatus('');
      this.hidePlayer(true);
      this.updateProgress(0);
      this.updateDuration(0);

      if(card.dataset.audioAutoload === 'true'){
        this.download();
      }
    }

    cleanup(){
      if(this.objectUrl){
        try{
          URL.revokeObjectURL(this.objectUrl);
        }catch(err){
          console.warn('Failed to revoke audio URL', err);
        }
        this.objectUrl = null;
      }
    }

    ensureStructure(){
      if(!this.card.classList.contains('audio-card')){
        this.card.classList.add('audio-card');
      }

      let header = this.card.querySelector('.audio-card__header');
      if(!header){
        header = document.createElement('div');
        header.className = 'audio-card__header';
        const infoWrap = document.createElement('div');
        infoWrap.className = 'audio-card__info';
        const title = document.createElement('h3');
        title.className = 'audio-card__title';
        title.dataset.role = 'audio-title';
        title.textContent = this.card.dataset.audioTitle || 'Audio-Datei';
        infoWrap.appendChild(title);
        const meta = document.createElement('div');
        meta.className = 'audio-card__meta';
        const duration = document.createElement('span');
        duration.dataset.role = 'audio-duration';
        duration.textContent = '0:00';
        meta.appendChild(duration);
        const size = document.createElement('span');
        size.dataset.role = 'audio-size';
        meta.appendChild(size);
        infoWrap.appendChild(meta);
        header.appendChild(infoWrap);
        this.card.insertAdjacentElement('afterbegin', header);
      }else{
        header.classList.add('audio-card__header');
        let infoWrap = header.querySelector('.audio-card__info');
        if(!infoWrap){
          infoWrap = document.createElement('div');
          infoWrap.className = 'audio-card__info';
          const titleEl = header.querySelector('[data-role="audio-title"]');
          if(titleEl){
            if(!titleEl.classList.contains('audio-card__title')){
              titleEl.classList.add('audio-card__title');
            }
            infoWrap.appendChild(titleEl);
          }else{
            const generatedTitle = document.createElement('h3');
            generatedTitle.className = 'audio-card__title';
            generatedTitle.dataset.role = 'audio-title';
            generatedTitle.textContent = this.card.dataset.audioTitle || 'Audio-Datei';
            infoWrap.appendChild(generatedTitle);
          }
          header.insertAdjacentElement('afterbegin', infoWrap);
        }else{
          const titleEl = infoWrap.querySelector('[data-role="audio-title"]');
          if(titleEl && !titleEl.classList.contains('audio-card__title')){
            titleEl.classList.add('audio-card__title');
          }
        }
        const metaEl = header.querySelector('.audio-card__meta');
        if(!metaEl){
          const meta = document.createElement('div');
          meta.className = 'audio-card__meta';
          const duration = document.createElement('span');
          duration.dataset.role = 'audio-duration';
          duration.textContent = '0:00';
          meta.appendChild(duration);
          const size = document.createElement('span');
          size.dataset.role = 'audio-size';
          meta.appendChild(size);
          (infoWrap || header).appendChild(meta);
        }
      }

      let actionBtn = this.card.querySelector('[data-role="audio-action"]');
      if(!actionBtn){
        actionBtn = document.createElement('button');
        actionBtn.type = 'button';
        actionBtn.dataset.role = 'audio-action';
        actionBtn.textContent = 'Herunterladen';
        const headerEl = this.card.querySelector('.audio-card__header');
        if(headerEl){
          headerEl.appendChild(actionBtn);
        }else{
          this.card.appendChild(actionBtn);
        }
      }
      actionBtn.classList.add('audio-card__action');

      let player = this.card.querySelector('[data-role="audio-player"]');
      if(!player){
        player = document.createElement('div');
        player.dataset.role = 'audio-player';
        player.hidden = true;
        player.innerHTML = `
          <div class="audio-card__progress">
            <span data-role="audio-current">0:00</span>
            <input type="range" min="0" max="0" value="0" step="0.01" data-role="audio-seek" />
            <span data-role="audio-duration">0:00</span>
          </div>
          <div class="audio-card__controls">
            <button type="button" class="audio-card__toggle" data-role="audio-toggle">Abspielen</button>
            <span class="audio-card__complete" data-role="audio-complete"></span>
          </div>
        `;
        this.card.appendChild(player);
      }
      player.classList.add('audio-card__player');

      let status = this.card.querySelector('[data-role="audio-status"]');
      if(!status){
        status = document.createElement('div');
        status.dataset.role = 'audio-status';
        status.className = 'audio-card__status';
        this.card.appendChild(status);
      }else{
        status.classList.add('audio-card__status');
      }

      const completeLabel = this.card.querySelector('[data-role="audio-complete"]');
      if(completeLabel){
        completeLabel.classList.add('audio-card__complete');
      }
    }

    bindEvents(){
      if(this.actionBtn){
        this.actionBtn.addEventListener('click', (ev)=>{
          ev.preventDefault();
          this.handleActionClick();
        });
      }

      if(this.toggleBtn){
        this.toggleBtn.addEventListener('click',()=>{
          if(this.audio.paused){
            this.audio.play().catch(err=>{
              console.warn('Audio playback failed', err);
              this.updateStatus('Wiedergabe nicht möglich.');
            });
          }else{
            this.audio.pause();
          }
        });
      }

      if(this.seekInput){
        const handleSeek = (value)=>{
          if(!this.metadataReady){
            return;
          }
          const numeric = Number(value);
          if(!Number.isFinite(numeric)){
            return;
          }
          const clamped = Math.min(Math.max(numeric, 0), this.duration || numeric);
          this.seekFromInput = true;
          try{
            this.audio.currentTime = clamped;
          }catch(err){
            console.warn('Unable to seek audio', err);
          }
          this.updateProgress(clamped);
          this.checkCompletion();
        };

        this.seekInput.addEventListener('input',()=>{
          handleSeek(this.seekInput.value);
        });

        this.seekInput.addEventListener('change',()=>{
          handleSeek(this.seekInput.value);
          this.seekFromInput = false;
        });
      }

      this.audio.addEventListener('loadedmetadata',()=>{
        this.onMetadataReady();
      });
      this.audio.addEventListener('durationchange',()=>{
        if(Number.isFinite(this.audio.duration) && this.audio.duration > 0){
          this.onMetadataReady();
        }
      });
      this.audio.addEventListener('timeupdate',()=>{
        if(this.seekFromInput){
          return;
        }
        this.updateProgress(this.audio.currentTime);
        this.checkCompletion();
      });
      this.audio.addEventListener('ended',()=>{
        this.updateProgress(this.duration);
        this.markComplete();
      });
      this.audio.addEventListener('play',()=>{
        this.setState(STATE.PLAYING);
        if(this.toggleBtn){
          this.toggleBtn.textContent = 'Pause';
        }
      });
      this.audio.addEventListener('pause',()=>{
        if(!this.audio.ended){
          this.setState(STATE.PAUSED);
        }
        if(this.toggleBtn){
          this.toggleBtn.textContent = 'Abspielen';
        }
      });
      this.audio.addEventListener('error',(ev)=>{
        console.error('Audio error', ev);
        this.setState(STATE.ERROR);
        this.updateStatus('Audio kann nicht geladen werden.');
      });
    }

    handleActionClick(){
      if(this.state === STATE.IDLE || this.state === STATE.ERROR){
        this.download();
        return;
      }
      const shouldShow = this.playerVisible ? false : true;
      this.togglePlayerVisibility(shouldShow);
      if(!shouldShow){
        return;
      }
      if(this.metadataReady && this.toggleBtn){
        this.toggleBtn.focus();
      }else if(!this.metadataReady){
        this.ensureMetadata();
      }
    }

    setState(next){
      if(this.state === next){
        return;
      }
      this.state = next;
      this.card.dataset.audioState = next;
      this.updateActionButton();
    }

    updateActionButton(){
      if(!this.actionBtn){
        return;
      }
      this.actionBtn.classList.remove('is-downloading','is-ready','is-complete');
      this.actionBtn.disabled = false;

      if(this.isComplete){
        this.actionBtn.textContent = 'Erledigt';
        this.actionBtn.classList.add('is-complete');
        return;
      }

      switch(this.state){
        case STATE.IDLE:
          this.actionBtn.textContent = 'Herunterladen';
          break;
        case STATE.DOWNLOADING:
          this.actionBtn.textContent = 'Lädt…';
          this.actionBtn.disabled = true;
          this.actionBtn.classList.add('is-downloading');
          break;
        case STATE.ERROR:
          this.actionBtn.textContent = 'Erneut versuchen';
          break;
        default:
          this.actionBtn.textContent = 'Öffnen';
          this.actionBtn.classList.add('is-ready');
          break;
      }
    }

    updateStatus(message){
      if(this.statusLabel){
        this.statusLabel.textContent = message || '';
      }
    }

    updateDuration(seconds){
      this.duration = Number(seconds) || 0;
      const formatted = formatTime(this.duration);
      this.durationLabels.forEach(el=>{
        el.textContent = formatted;
      });
      if(this.seekInput){
        this.seekInput.max = this.duration ? String(this.duration) : '0';
        this.seekInput.disabled = !this.duration;
      }
      if(this.toggleBtn){
        this.toggleBtn.disabled = !this.duration;
        if(!this.duration){
          this.toggleBtn.textContent = 'Abspielen';
        }
      }
    }

    updateProgress(seconds){
      const value = Number(seconds) || 0;
      if(this.currentLabel){
        this.currentLabel.textContent = formatTime(value);
      }
      if(this.seekInput && !this.seekFromInput){
        this.seekInput.value = String(Math.min(Math.max(value, 0), this.duration || value));
      }
    }

    hidePlayer(force){
      if(!this.playerWrap){
        return;
      }
      this.playerVisible = !force;
      this.playerWrap.hidden = !!force;
    }

    togglePlayerVisibility(show){
      if(!this.playerWrap){
        return;
      }
      const shouldShow = show === undefined ? !this.playerVisible : !!show;
      this.playerWrap.hidden = !shouldShow;
      this.playerVisible = shouldShow;
      if(!shouldShow && this.audio && !this.audio.paused){
        try{
          this.audio.pause();
        }catch{}
      }
    }

    ensureMetadata(){
      if(this.metadataReady){
        return;
      }
      if(this.metadataAttempts >= METADATA_RETRY_LIMIT){
        if(!this.duration){
          this.updateStatus('Audiometadaten konnten nicht geladen werden. Bitte erneut öffnen.');
        }
        return;
      }
      this.metadataAttempts += 1;
      try{
        this.audio.load();
        this.audio.currentTime = this.audio.currentTime || 0;
      }catch(err){
        console.warn('Metadata reload failed', err);
      }
      clearTimeout(this.metadataTimer);
      this.metadataTimer = setTimeout(()=>{
        if(!this.metadataReady){
          this.ensureMetadata();
        }
      }, METADATA_RETRY_DELAY);
    }

    onMetadataReady(){
      clearTimeout(this.metadataTimer);
      this.metadataReady = true;
      const duration = Number(this.audio.duration);
      if(Number.isFinite(duration) && duration > 0){
        this.updateDuration(duration);
        this.updateStatus('Audio bereit.');
        this.setState(STATE.READY);
      }
    }

    download(){
      if(this.downloadPromise){
        return;
      }
      if(!this.source){
        this.updateStatus('Keine Audio-Quelle verfügbar.');
        return;
      }
      if(this.isComplete){
        this.isComplete = false;
        this.card.dataset.audioComplete = 'false';
        if(this.completionLabel){
          this.completionLabel.textContent = '';
        }
      }
      this.setState(STATE.DOWNLOADING);
      this.updateStatus('Lade Audio-Datei…');

      const controller = new AbortController();
      this.downloadPromise = fetch(this.source, {signal: controller.signal})
        .then(response=>{
          if(!response.ok){
            throw new Error(`Download fehlgeschlagen: ${response.status}`);
          }
          return response.blob().then(blob=>({blob, response}));
        })
        .then(({blob, response})=>{
          if(this.objectUrl){
            URL.revokeObjectURL(this.objectUrl);
          }
          this.objectUrl = URL.createObjectURL(blob);
          this.audio.src = this.objectUrl;
          this.metadataReady = false;
          this.metadataAttempts = 0;
          this.ensureMetadata();
          const size = blob.size || Number(response.headers.get('content-length'));
          if(this.sizeLabel){
            this.sizeLabel.textContent = size ? formatBytes(size) : '';
          }
          this.setState(STATE.READY);
          this.updateStatus('Datei erfolgreich geladen.');
          return blob;
        })
        .catch(error=>{
          if(error.name === 'AbortError'){
            return;
          }
          console.error('Audio download failed', error);
          this.setState(STATE.ERROR);
          this.updateStatus('Download fehlgeschlagen. Bitte erneut versuchen.');
        })
        .finally(()=>{
          this.downloadPromise = null;
        });

      this.card.addEventListener('audio:cancel-download',()=>{
        controller.abort();
      },{once:true});
    }

    checkCompletion(){
      if(this.isComplete || !this.metadataReady || !this.duration){
        return;
      }
      const ratio = this.audio.currentTime / this.duration;
      if(ratio >= COMPLETION_THRESHOLD){
        this.markComplete();
      }
    }

    markComplete(){
      if(this.isComplete){
        return;
      }
      this.isComplete = true;
      this.card.dataset.audioComplete = 'true';
      if(this.completionLabel){
        this.completionLabel.textContent = '✓ Gehört';
      }
      this.updateStatus('Audio vollständig angehört.');
      this.updateActionButton();
    }
  }

  function initAudioCards(){
    const cards = Array.from(document.querySelectorAll('[data-audio-item], .audio-card[data-audio-src]'));
    cards.forEach(card=>{
      if(card.__audioController){
        return;
      }
      card.__audioController = new AudioCardController(card);
      const library = card.closest('.audio-library');
      if(library && library.hasAttribute('hidden')){
        library.removeAttribute('hidden');
      }
    });
  }

  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', initAudioCards);
  }else{
    initAudioCards();
  }

  window.HaasAudioPlayer = {
    refresh: initAudioCards,
    createCard({src, title, listId='audioList', description}={}){
      if(!src){
        return null;
      }
      const list = document.getElementById(listId);
      if(!list){
        return null;
      }
      const card = document.createElement('article');
      card.className = 'audio-card';
      card.dataset.audioItem = 'true';
      card.dataset.audioSrc = src;
      if(title){
        card.dataset.audioTitle = title;
      }
      if(description){
        const desc = document.createElement('p');
        desc.className = 'audio-card__description';
        desc.textContent = description;
        card.appendChild(desc);
      }
      list.appendChild(card);
      initAudioCards();
      return card;
    }
  };
})();
