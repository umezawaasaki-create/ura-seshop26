// URL of the deployed Apps Script Web App (ends with /exec). This is the
// only thing you need to edit when re-deploying the API to a new URL.
var SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbwjlw9mCYXlsulvgTNmK1FKV4jKo52Z8qTVQBYd6V7kz6iqMHPw7iVwAthYyOtsOeFeEg/exec';

/**
 * Read-only call (?action=...), GET request.
 */
function apiGet_(action) {
  var url = SCRIPT_URL + '?action=' + encodeURIComponent(action);
  return fetch(url).then(function (res) { return res.json(); }).then(function (data) {
    if (data && data.error) throw new Error(data.error);
    return data;
  });
}

/**
 * Call that writes data, POST request. Sent as text/plain (not
 * application/json) on purpose - that keeps it a CORS "simple request" so
 * the browser skips the preflight OPTIONS call, which Apps Script web apps
 * can't answer.
 */
function apiPost_(action, payload) {
  var body = Object.assign({ action: action }, payload || {});
  return fetch(SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  }).then(function (res) { return res.json(); }).then(function (data) {
    if (data && data.error) throw new Error(data.error);
    return data;
  });
}

var teamsById = {};
  var matchesById = {};
  var matchList = [];
  var currentMatchId = null;
  var currentIconTeamId = null;

  var COLUMNS = [
    { side: 'left', label: 'ROUND 16', ids: ['R16-1', 'R16-2', 'R16-3', 'R16-4'] },
    { side: 'left', label: '準々決勝', ids: ['QF-1', 'QF-2'] },
    { side: 'left', label: '準決勝', ids: ['SF-1'] },
    { side: 'center', label: '', ids: ['F', '3RD'] },
    { side: 'right', label: '準決勝', ids: ['SF-2'] },
    { side: 'right', label: '準々決勝', ids: ['QF-3', 'QF-4'] },
    { side: 'right', label: 'ROUND 16', ids: ['R16-5', 'R16-6', 'R16-7', 'R16-8'] }
  ];

  var ROUND_LABELS = {
    R16: 'ラウンド16', QF: '準々決勝', SF: '準決勝', F: '決勝', '3RD': '3位決定戦'
  };

  var AUTO_REFRESH_INTERVAL_MS = 20000;

  window.onload = function () {
    loadData();
    window.addEventListener('resize', debounce(function () {
      syncMobileBracketLayout_();
      drawConnectors();
    }, 150));
    startAutoRefresh_();
  };

  function loadData() {
    apiGet_('getBracketData').then(renderAll).catch(handleError);
  }

  /**
   * Polls the spreadsheet in the background so other people's schedule/score
   * updates show up without a manual reload. Skips a poll while the tab is
   * hidden (saves quota/battery) and while the edit modal is open (so it
   * doesn't yank the bracket around under someone mid-edit) - it just picks
   * up on the next tick instead.
   */
  function startAutoRefresh_() {
    setInterval(function () {
      if (document.hidden) return;
      if (!document.getElementById('modal-backdrop').hidden) return;
      loadData();
    }, AUTO_REFRESH_INTERVAL_MS);

    document.addEventListener('visibilitychange', function () {
      if (!document.hidden) loadData();
    });
  }

  function renderAll(data) {
    teamsById = {};
    (data.teams || []).forEach(function (t) { teamsById[t.team_id] = t; });
    matchesById = {};
    matchList = data.matches || [];
    matchList.forEach(function (m) { matchesById[m.match_id] = m; });

    renderColumns();
    requestAnimationFrame(function () {
      syncMobileBracketLayout_();
      drawConnectors();
    });
  }

  // ---------------------------------------------------------------------
  // Mobile view switching + zoom
  // ---------------------------------------------------------------------

  var mobileZoom = 1;
  var naturalBracketSize = null;

  function isMobileLayout() {
    return window.innerWidth <= 760;
  }

  function sizeBracketWrapForMobile_() {
    var wrap = document.getElementById('bracket-wrap');
    if (!wrap) return;
    if (!isMobileLayout()) {
      wrap.style.height = '';
      return;
    }
    var top = wrap.getBoundingClientRect().top;
    var h = window.innerHeight - top - 8;
    wrap.style.height = Math.max(200, h) + 'px';
  }

  /**
   * Keeps the bracket's mobile zoom/sizing in sync with the current viewport.
   * There's no manual zoom control - it's always fit-to-screen. On desktop
   * widths this clears any leftover mobile transform/sizing (e.g. after
   * rotating a device or resizing a browser window past the mobile
   * breakpoint) so the graph renders at its normal, untransformed size.
   */
  function syncMobileBracketLayout_() {
    sizeBracketWrapForMobile_();

    if (!isMobileLayout()) {
      fitBracketWidthToShell_();
      return;
    }

    // renderColumns() rebuilds every .match-card from scratch on each
    // refresh, which wipes the inline counter-rotation style those elements
    // carry on mobile - fitBracketToScreen() (via applyMobileZoom_) re-applies
    // it every time alongside the fit-to-screen scale.
    fitBracketToScreen();
  }

  /**
   * Desktop/wide-layout safety net: the bracket must never spill out past
   * the translucent panel. If its natural width is wider than the visible
   * bracket-wrap area (e.g. a narrower browser window, or the cards grew),
   * scale the whole bracket down just enough to fit - no horizontal
   * scrolling, no cards cut off at the panel edge. Full size (no scaling)
   * whenever it already fits.
   */
  function fitBracketWidthToShell_() {
    var wrap = document.getElementById('bracket-wrap');
    var bracket = document.getElementById('bracket');
    var scaleWrap = document.getElementById('bracket-scale-wrap');
    if (!wrap || !bracket || !scaleWrap) return;

    setCounterRotation_(false);
    measureNaturalBracketSize_();
    if (!naturalBracketSize.width) return;

    var availW = wrap.clientWidth - 32;
    var scale = availW > 0 ? Math.min(1, availW / naturalBracketSize.width) : 1;

    if (scale >= 0.999) {
      bracket.style.transform = '';
      scaleWrap.style.width = '';
      scaleWrap.style.height = '';
      mobileZoom = 1;
      return;
    }

    mobileZoom = scale;
    bracket.style.transform = 'scale(' + scale + ')';
    scaleWrap.style.width = (naturalBracketSize.width * scale) + 'px';
    scaleWrap.style.height = (naturalBracketSize.height * scale) + 'px';
  }

  function measureNaturalBracketSize_() {
    var bracket = document.getElementById('bracket');
    naturalBracketSize = { width: bracket.offsetWidth, height: bracket.offsetHeight };
  }

  /**
   * Rotating the whole bracket 90deg would rotate every label/badge/logo
   * with it, sideways and unreadable. Counter-rotating each card/label/the
   * center trophy by -90deg cancels that back out for their own content -
   * net rotation zero, so they render upright - while their position still
   * follows wherever the bracket's own rotation placed them. The connector
   * lines are not counter-rotated: they're meant to visually follow the
   * rotated layout.
   */
  function setCounterRotation_(on) {
    var els = document.querySelectorAll('.match-card, .round-label, .center-deco');
    els.forEach(function (el) { el.style.transform = on ? 'rotate(-90deg)' : ''; });
  }

  function applyMobileZoom_(scale) {
    if (!naturalBracketSize) measureNaturalBracketSize_();
    mobileZoom = Math.max(0.15, Math.min(1.5, scale));

    var bracket = document.getElementById('bracket');
    var scaleWrap = document.getElementById('bracket-scale-wrap');

    if (isMobileLayout()) {
      bracket.style.transform = 'rotate(90deg) scale(' + mobileZoom + ')';
      scaleWrap.style.width = (naturalBracketSize.height * mobileZoom) + 'px';
      scaleWrap.style.height = (naturalBracketSize.width * mobileZoom) + 'px';
      setCounterRotation_(true);
    } else {
      bracket.style.transform = 'scale(' + mobileZoom + ')';
      scaleWrap.style.width = (naturalBracketSize.width * mobileZoom) + 'px';
      scaleWrap.style.height = (naturalBracketSize.height * mobileZoom) + 'px';
      setCounterRotation_(false);
    }
  }

  /**
   * The bracket is rotated 90deg on mobile (see applyMobileZoom_), so its
   * natural width ends up running along the screen's vertical (scrollable)
   * axis and its natural height along the horizontal one. Scale is picked to
   * fill the full screen width - not capped by available height too - since
   * .bracket-wrap already scrolls vertically; capping by both would leave
   * unused space on the sides whenever the height constraint was tighter.
   */
  function fitBracketToScreen() {
    if (!isMobileLayout()) return;
    measureNaturalBracketSize_();
    var wrap = document.getElementById('bracket-wrap');
    if (!wrap || !naturalBracketSize.width || !naturalBracketSize.height) return;

    var availW = wrap.clientWidth - 16;
    if (availW <= 0) return;

    var fit = Math.min(availW / naturalBracketSize.height, 1);
    applyMobileZoom_(fit);
  }

  function renderColumns() {
    var bracketEl = document.getElementById('bracket');
    var html = '<svg id="connectors"></svg>';

    COLUMNS.forEach(function (col) {
      if (col.side === 'center') {
        html += '<div class="col-outer center">';
        html += '<div class="bracket-col center">';
        html += '<div><div class="round-label center-round-label"><img class="final-label-img" src="final.svg" alt="FINAL"></div>' + matchCardHtml(matchesById['F']) + '</div>';
        html += '<div class="center-deco">' + trophyHtml() + '</div>';
        html += '<div><div class="round-label center-round-label">3位決定戦</div>' + matchCardHtml(matchesById['3RD']) + '</div>';
        html += '</div></div>';
      } else {
        html += '<div class="col-outer">';
        html += '<div class="round-label">' + col.label + '</div>';
        html += '<div class="bracket-col">';
        col.ids.forEach(function (id) {
          html += matchCardHtml(matchesById[id]);
        });
        html += '</div></div>';
      }
    });

    bracketEl.innerHTML = html;
  }

  function matchCardHtml(match) {
    if (!match) return '';
    var t1 = teamsById[match.team1_id];
    var t2 = teamsById[match.team2_id];
    var hasTeams = !!(match.team1_id && match.team2_id);

    var dateTimeParts = [];
    if (match.date) dateTimeParts.push(formatDateLabel(match.date));
    if (match.time) dateTimeParts.push(match.time);

    var metaLines = '';
    if (dateTimeParts.length) {
      metaLines += '<div class="meta-line">' + escapeHtml(dateTimeParts.join(' ')) + '</div>';
    }
    if (match.venue) {
      metaLines += '<div class="meta-line venue">📍' + escapeHtml(match.venue) + '</div>';
    }
    if (match.win_type === 'fusen1' || match.win_type === 'fusen2') {
      metaLines += '<div class="meta-line"><span class="win-tag">不戦勝</span></div>';
    }

    return '' +
      '<div class="match-card ' + (hasTeams ? '' : 'tbd') + '" data-match-id="' + match.match_id + '" onclick="openModal(\'' + match.match_id + '\')">' +
      teamRowHtml(t1, match.team1_id, match.score1, match.winner_id) +
      teamRowHtml(t2, match.team2_id, match.score2, match.winner_id) +
      '<div class="meta-row">' + metaLines + '</div>' +
      '</div>';
  }

  function teamRowHtml(team, teamId, score, winnerId) {
    if (!team) {
      return '<div class="team-row"><span class="badge" style="background:#999">?</span><span class="team-name">未定</span></div>';
    }
    var isWinner = winnerId && winnerId === teamId;
    var isLoser = winnerId && winnerId !== teamId;
    var scoreHtml = (score !== '' && score !== undefined && score !== null) ? '<span class="team-score">' + escapeHtml(String(score)) + '</span>' : '';
    return '' +
      '<div class="team-row ' + (isWinner ? 'is-winner' : '') + ' ' + (isLoser ? 'is-loser' : '') + '">' +
      badgeHtml(team) +
      '<span class="team-name">' + teamNameHtml_(team.name) + '</span>' +
      scoreHtml +
      '</div>';
  }

  // Combined team names like "二子玉川・瀬田" break onto a second line at the
  // "・" instead of wrapping wherever happens to fit, which reads awkwardly
  // in the narrow mobile cards.
  function teamNameHtml_(name) {
    return escapeHtml(name).replace(/・/g, '・<br>');
  }

  /**
   * Renders a team's logo. If a logo_url has been set on the Teams sheet,
   * that image is used; otherwise this falls back to the plain colored
   * circle + initial placeholder, so teams keep working with no image set.
   */
  function badgeHtml(team) {
    if (team.logo_url) {
      return '<img class="badge-img" src="' + escapeHtml(team.logo_url) + '" alt="' + escapeHtml(team.name) + '">';
    }
    return '<span class="badge" style="background:' + escapeHtml(team.color || '#666') + '">' + escapeHtml(team.initial || '') + '</span>';
  }

  /**
   * The tournament emblem shown between the final and 3rd-place match cards:
   * a maroon badge with a white "P" letterform, a trophy silhouette, and the
   * "世小P" / "URA SSP 2026" wordmarks. Shared by the desktop bracket's
   * center column and the mobile round list.
   */
  var LOGO_DATA_URI = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMMAAAFTCAYAAACEbrhdAAAACXBIWXMAAAsTAAALEwEAmpwYAAAFxGlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgOS4wLWMwMDEgNzkuMTRlY2I0MmYyYywgMjAyMy8wMS8xMy0xMjoyNTo0NCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0RXZ0PSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VFdmVudCMiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIDI0LjIgKE1hY2ludG9zaCkiIHhtcDpDcmVhdGVEYXRlPSIyMDI2LTA3LTE1VDIwOjA4OjIyKzA5OjAwIiB4bXA6TW9kaWZ5RGF0ZT0iMjAyNi0wNy0xNVQyMTo0MDowNyswOTowMCIgeG1wOk1ldGFkYXRhRGF0ZT0iMjAyNi0wNy0xNVQyMTo0MDowNyswOTowMCIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6OTBkOTkzNjItMmU5Ny00ZjFmLWE2NDctN2NhYzI5MDliZDQzIiB4bXBNTTpEb2N1bWVudElEPSJ4bXAuZGlkOmQwYWNmNDhkLWY3MjktNDU4OS1iM2E1LTFhYWEyN2I1MDhiNiIgeG1wTU06T3JpZ2luYWxEb2N1bWVudElEPSJ4bXAuZGlkOmQwYWNmNDhkLWY3MjktNDU4OS1iM2E1LTFhYWEyN2I1MDhiNiI+IDx4bXBNTTpIaXN0b3J5PiA8cmRmOlNlcT4gPHJkZjpsaSBzdEV2dDphY3Rpb249ImNyZWF0ZWQiIHN0RXZ0Omluc3RhbmNlSUQ9InhtcC5paWQ6ZDBhY2Y0OGQtZjcyOS00NTg5LWIzYTUtMWFhYTI3YjUwOGI2IiBzdEV2dDp3aGVuPSIyMDI2LTA3LTE1VDIwOjA4OjIyKzA5OjAwIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgMjQuMiAoTWFjaW50b3NoKSIvPiA8cmRmOmxpIHN0RXZ0OmFjdGlvbj0ic2F2ZWQiIHN0RXZ0Omluc3RhbmNlSUQ9InhtcC5paWQ6OTBkOTkzNjItMmU5Ny00ZjFmLWE2NDctN2NhYzI5MDliZDQzIiBzdEV2dDp3aGVuPSIyMDI2LTA3LTE1VDIxOjQwOjA3KzA5OjAwIiBzdEV2dDpzb2Z0d2FyZUFnZW50PSJBZG9iZSBQaG90b3Nob3AgMjQuMiAoTWFjaW50b3NoKSIgc3RFdnQ6Y2hhbmdlZD0iLyIvPiA8L3JkZjpTZXE+IDwveG1wTU06SGlzdG9yeT4gPC9yZGY6RGVzY3JpcHRpb24+IDwvcmRmOlJERj4gPC94OnhtcG1ldGE+IDw/eHBhY2tldCBlbmQ9InIiPz6ucfdoAACCMElEQVR4nO29d5wkR333/66q7pnZ3cs5KGcJJJQAEQwiiCCSZaJtsAFjg3F+fn6SjdNjYxsHkh/A+MEG22QwQSCEhISEBMo5h7vTnS6nzTsz3V1V398f1d0zu7eXdvZud0/zuVffzM5Md1VX17fqm79KROiiiy5Az3QHuuhitqBLDF10kaNLDF10kaNLDF10kaNLDF10kaNLDF10kaNLDF10kaNLDF10kaNLDF10kaNLDF10kaNLDF10kSOa6Q7MFtzy4Q+JLFzC2uNPZPHiRVjrSLIErTVRFIEC7w/sx+UB1FHp7j4wOqIaVWkmCUqB9o5IPFseephn/f4fzlCv5ha6xJDjhR/8TejtA6VBBLQOh8o3TxEQf4ArzPAm6zU4wwIRMAokg7SBuJnt1lxClxgKGAOmikSV8iNBEAmHUqCNyT+fDGrC69GFRoEFEJzPMJUYIk8SdTeFQ0WXGAooAxKhxICAtw6tNZiw4jvnQAIxqPK/NsywJ7xXgteeDIeOFEJGFHnSqOuif6joEkMBHYXDAwp0FCY+Lvxt8l1hv5jpBVgplDKApk5CFYXBMKa6OpJDhZpycM/YDiHWYCrgNBAexIxhn8kordVaAFGt3ynVOrTKPztw/2f7+qoE8EF0SFW4kxhQTsBZcA7EgRaIfC4bReEErwEB7Zj9d7ofiANvGf7pTSy47I1TWpqmvjNUYmwcAVW0MSgMKsyqKV+yE+RzYRw0eW/aiaKNtXf5OY7W5DlctLepy0ZmAirM51x+Lp+EVqDjIBN5DwZEKwQd5AyhNS5qKiMwSyAWXEozilgwxUtMnRh8hPYVlK6AaFQ562ZmMhyQDBXjCQFAFEa1CCYch9//9r1kRjklJeUuFyGtvijBI2iT31cu8GgIlFyMSzk+c3RnUAqiiERPnTuZOjGIQUvOWhTjp4pRnRlMbF3aPt/ny/zpKwKD1yLkw+v/TIsKLQgoD0radujQO5cTRxAf9qP1mj03MkUoBIPVB5HtDoCpE4MKC8u+dDhzo1q03M66CG3skmr7cJ+5MNdnQzsht+91CsP4XaG1azAJbczNcRAUgqD91PvfgTbJ51Ib4ybXDHLMYXHMWZ92giiI9uDPezpknhkcAVX0XyMofL5UjV8rWwKUqLk58fcHQaFkJtgk1c5WqPJlxoa3rTtK5UZY9t0lJu/f/r85lGYZd/ZMjUAx5YNwXCgHDLQGodCkieQKBJ+fqfKv5zZxGIGog52hM12ozOLBk2KNPPhNzuK7OEzs5273pdh9fnBMjIEUSpCpoTNtkopy1Z1v26JnCB3JhVPv9+yaRKr837S97kMfOXtkcq2Cajt3TkMFS/xU0YE2SZfuCaXsMNMDegw8z+lCS0vGfsflmCGCNnRCDF1bfRdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSQxdd5OgSwzRCvOwT9Vmkp5w0UwH7fiYieO/L1/bvRQRnXcj4l3/nvSdN02m9j2cquhn1phMKvPgQHq4VSoWjQJGwrZ04FK3feO/RWqOUCkTkpbVcqfBbnae7LAjDRIbIROX7LqaOLjFMIyad/BI+D/H6qvxdO4qJLyKIEqy1aK0xkUG8kNmsJJA4jinSXYoIzrmDp77s4pDQJYZphHgpJz607QB5ziZrLVEUhd8IOB/yxRttUHkGAy+eOIrx4hEvWGepVNoyg4tgM4vSCmNMSRQTCayLw0eXGKYRSueT3DmUUiGLdxviOA6T2VqUUuNW9EJOMNqAAo0O2fFUYIGUUiXrZSKDsw5H2BUUqmSxupg6usQwXchXda10OcnbhWCV59FRqHICt6N9IheskTGmnPztn8dxTBRHpQyCAt3Ntt0xusQwXVCgJGR1U9KSEQo2pl6v89hjj7F+/Xo2bNjA6MgocSUmjmIym2GtJY5jsixj8eLFXPL8S6hUK1x88cVALlcQ2KEszYjiKOwIzpdtdVmlztAlhmmE0q20gmNjYzzyyCP84Ac/4Nvf/jb333//PgJ0sbK3T+IoikjTtGSzoijizDPP5LLLLuOKK67gkudfQlSJSuLw4oPM0SWEjjH1+gyNhmB6iiIAtFJ2HeOQUN7KOz+pKnP79u189rOf5ctf/jJPPfVUqe2x1nbctDGGk046ife+9738yrt+hbXHrUVE0FqXQnShYYpMhPe+tWM8E2hFYPNPruP4S185xfSIIlM76nWRRES8iIgVESfPBHjvy1fvvKRpKt57Wb9+vfzGb/yGLFiwQLTWEkWRGGOkUqmI1lriOG7PxzmlQyklWmtRSsmyZcvkfe97n9xzzz3ivZcsy8RmVpJm0tZZEeecNJvNst/HNLzI0zf8SGSKc7pLDIcB51w5+Qvs2b1H/uEf/kFWrVolSikxxpQTtn0Cd0oI7QTRfs1Vq1bJh//qw9JoNFodfQbM+0nRJYajB+dcufImSSK33367XHbZZeNW7CiKBBi3GxSE0cmhtZZKpVJeS2stxhgBJI5jee5znys//OEPJU3T0M8kKfvcbDTFWjuTQ3d00CExPAOY/OlDoeYcGRnha1/7Ou985zu56aabSrVoYRyLosCvZ1nLctypgFv4IEmbjOBcMNplWcZdd93FO9/5Tv7sz/6MZqNJJa6Ulu1qrdq1Uh8CusRwqMh30Xq9zje/+U1+//d/j3Xr1pEkCc45RIQ0TfHel0a3OI5LQpGpKiraoLWmVquVQnG1Wm3ZIoxheHiYT3/607znve9h/Yb1pZW76F8XB0ZXtXqoULBj+w4++tGP8slPfpIsy4CwCxhjyl2gIASRlo/RRFXqlJrPr9FsNsu/kyQpCcF7j7WWLMu48sor2blzJ//2r//GKaecMkkJry4mQ3dnOERs2riJD33oQ3zsYx8rCQGC60WSJHjvS5eIgpUBSvVmpyuztFuz21ivghDa3TGazSa33HIL7/2197Ju/bpw3jNBtdohusSQo50HFxGyLCsn8KOPPsr7P/B+Pv/5z4+PMWD8at8+YQtfo/bPprOvxau1dp82ITgF3nHHHbz//e9n/br149zGi98WLN0xwULlt6A7uJUuMeRQSqFQNJvNkt9XSrFjxw7+8z//k2uvvXamu3hQFBbrYsdoNBrccsst/O8/+t+MjY1hnS2JXmuNVjr4SBU6q7mODu+hSwxt0CYIqN55ms0mg4ODfOUrX+Fv//Zv54RHaLHSF6yTUgprLT/+8Y/50z/9U7TWZGnWioXwDm30sWOhdh7npm7pn/1P+GhBwLvAOiitiKOY2267jT/+4z/GGFOyULMZxc7QzjpZaxkcHOR73/seX/ziF4krMd758n6mm4WbOYTtrZNFq0sM7VCUAvBDDz/Ehz/8Yay1cyaarJBRCic/pVSpil2/fj2f+tSneOSRR0KcRBRhdGuHmPtygwJjOnJl7xJDgTzWIMsyRoZH+MlPfsKtt96KtfaICMFHAoVAXUxsYwzNZhOtNZVKhfXr1/P5f/t8iK0mOBwaY0qimPPwnjRpTvn0LjG0I48LeOLJJ/jYxz5W2gwqlcqsd5Fuj6wrdrcizLRQ//b393PNtddw3fXXhSwbMC3etLMCImT1Mfbu3TPlS3SJoQ1KKfr7+7n66qvZsmVLKYhmWTbrd4aCRWpn5wojYKFhUkqxbt06/u3f/q1ludYmCNJzQEFwQKhg8xkbHZ3yJeb4CEwvvPds2bKFq666CufcrCeAiSjUpu15l9ptJoV/07p167j3vnsDAZHHXR8DUAjWZgf/4X7QJYYCAmmasmnTJu644w6UUkRRFFK2zAHh+VBQ2B+efvppvvGNb7Qs5seEkSHUeu+Eme0SQwEF27Zu42tf+1rJf3vvqVQqJasxl1EQQpZl9Pf3c/vtt9NsNqfFo3ZWQDziPD3V6pQv0SWGNvQP9HP33XcDrXQvSZJMW9jmTKJgkwqD2+7du7n77ruPDUKAPEWPJUkaU75ElxhyJM2E9evXs27dulLYLAhiLhjcDgXFLqCUYmBggAcffLC0Ws95aIVzlm3btk/9EtPYnTmNgYGBfVbKdlnhWGCTCsLWWjMyMsKTTzy5jwPfXEaWZYyMDE/5/Ln9hKcRjUaDrVu3AkFP75wbl/lurrNJ0CJuay0jIyPc/8D9x4xyABFwHt/BLtfBziCghOD0qJjrnl5Dw0M8/PDDpR+S1rp0aTgWCKEwwhUu24XKdfu27ccGm4SgfErWnAmZQQHKI4BD41DTk/6ho0P2+12BkAQhOOUhQF5TYWR4mCeeeCL/zXh9/bGEwpvVGMPY2Bg7d+1Ea43NbEkkzrny/pMkOWp928eTvHyAkz3JCfAeOzzA7/zrl6e8KncoM4TOFenD1Awf7X1qh5rwXrzkrJAFBG+DUcrux9J8rGhcCq/WYtIXKS9FJKSrzO/TGFOqlqvV6lGRKQQ4NBOn7IcshMh3tsNNnRjyHikBI/mEm+GtoUUYLcatTNLiARuGW6ughjNGg1YkSZN169aR5a7PhZ1hzrsoTECx0/X29paRboVc1F4BaFyyZML0m1FWaj9r0USC8B1WMJq6AF1uXcX7mV49D6YVEZx4JAsVbooKO2makNiMkbGR8pftAfwFmzQdccwzjcLWkCQJIkKj0aC/v7/UNFUqlVYeIS9lsZTCtX2Gen1oP/Oe5thYRy11pk0qGPACapZMFpnELq8URkeBCPL3ToSop0r/yDBpHiE1GSEcC2j3ai3sJ/V6nS2bt4RdIUnRJrh6FyxS4bvUXmrr6HZ6MmYobPL7kGaSMbSnn7UdNNcZMUxkxpk5lbVqf3eQ55bYjIHBAb7//e/z5JNP8sHf+i2qtZ5yhTzWCKFAu2drYYCzzrJl8xY+/4XPs2LFSt7//t8YX0hFWuzSkcS+aasnJ4Ti4e5DEGnCUAfu29AJMaj8v/YaZjPOKk2CYkyL3EECtajCts1b+dg/fgytNc857zlkzaA1OZbYonYU91LYUAqV8djYGPfedy9/8Rd/wSWXXMIv/MIVLFu6LGQYF8qd4ehUBjrYeKv9Ctm+2WBg546OWu/g7nTpJih5goWZ1yS1YaLKwQNeQioRD8MDQzz56GP079rDw/c/iGbfHaHd1nCsEEZ7mvrge5WyceMmIBget2zeEmQFackKQmexxYeD/S2nMoEQ9PgvyZpNBnft7qjtDgJGNRiDBSxhrokHcbPjwAOu/ZBcd6fBQwWDdtAcHuX+O+4hVgbT9sDb+eZjRbU6McfTvHnzmD9/HjfffFP5XRSHGnLjtEcHUfEfOYxf7oRQ8FEI5by897mdSGiMjPJL//yFjh5UB9qk3ODWxlYozWwQGgJkwntRYeC8BN1qjnq9wd69e0HA5/aHdue8YgUtsk7MVbS7pRcE3tvbi1KKW2+9dZxwXbBI5bn66CwGaj9/tTjdIMtoQOWvQL7wda767YgYlAKDwnuHdR6ieEY1rO3zvwj00Dn/phSgFeKD37vTUOnrJU0zNm3bxjXXX8/iJUsYHh4aRwha61IXfyygfacTEZIkYdu2beXf5e9yWeHo7owHbyekehO0UuWugLOMDPSzqMPWO9AmKfAO8R6tFHFksGJnbGOAsPgXxjchEELBZxYGpjiKMXFE34L5nHHO2dxzzz3sHNjL3ffew4qVKxgcHNhnwrS/zlW0p88siHxkZIRHH3203PVqtRpLliwZt3uUJbIOQUt3JOHze9BK8CIYgCI3VLPBWAfeqgU60CYJeE9cdkqIZ4MHpGpjcNsentGaKIqx1mEBMUIza4D24B2DI4PEUXjwhbt24dhWCNJzmU2CFstXWJhHR0e5++67sdZijGHevHmsWbNmfBp9YVbITQpBxBMWYQ9KByFVKdJGkx3btnFWh21MnRi8D50phE5nw98zjsmkvfAglVLEJvgkza9VefjeB5SKtQR2ocnOXTvLM9pX0WPJ7tBe7KTZbJb3GUUR8+fPLwuulFDM+K6QPz0UCq0UUvDAOYEm9TqX/tbvddzDqRNDvQHzDCRpIAJjZjZUSMr/2j5oe4iFo5ILW+yqvgUAxDZ3/EotKRaV6+Hbo8KOmWgwKJMqF+lvil2wt7eXU089tawrPfGcmUYQ+VT5vt3zoT46wuJpaGPqxLBsxcyPUAdYeNzxAMwDUkA0NEWNc3FuXzXnOotUoPBNKjxYiyIrCxcu5Nxzz2vVlRCPVjq4r8B+S/0eUeTrWWva+9yopXI9PiAe3xjZ7yUOB8eWW+YUsECRC2PjheX2Wg3HCiG0W9fbDYkiwoIFCzj99NPw4kNcSG50K36rzZGfKi0HZD/ugyCRCiI2lxcM6BinIhCP7e/M8lzgGU8MdYEMyOb0PnfoKPJATTQo9vX1cdKJJwGBOIw2pfAMR8c/CfJJP2lMikJrlQvRgSa00di9ezj5NW+els5NnU0a2iREGipVsAKmBlJhZl25JcgFyrW0SoWruahghbYClR5Asenpzbz2V3+V//jq1yBLW789BlGoSYsItnY5oKenh1NOOZXVa1YDlLvC0VYc7DNz8qaNKjkiVBSFx6zCI87q9WkL5J/6dXp7IY6wGFysEaepmZ6Zn0tK8OVmG2r76qJT1geeSFdwAvNXrOU5z3sxfOkbENUgnXr87GzHxJpwhTVaa82KFSt47nMvDgQiucVZ2nYIWka4I41gI5rQjgTVqsvtC05Aq7DQZaNj9ExT21Nnk5whs4qMCmNWY0xPzuvN1EHpdiFoHAZPhCUiIybzhswpiKqAwqBY0tPLJec8h764D9Kp5+icSyi0R+2q40WLFvH85z8/sE75wiFIq6KPHCWNksA++4O0PjfGYL3DKY/zGbiMPdu3TVvzHcgMMXHUR0yV3qg3eBUK+HxFPtqHiAeRYOoIEldJJg5Aa6JKJQjKNtdEaFi6aB7Pu+BceuZ4XqRDRbEbtOdRWrhwIWefdXZLk+RaOwgwKQ9/ZKHGvRTNiwTPVa8EYxR2eJCNjz02ba1OfQZYBU2Xb6OKKAbpoGpKp1AqOOEpJRil8gQFanw4gxJUpFAG8Ba0YuHy+Vzykudy4y0/nrG+Hw20ywwFIRhj6O3t5YILLmDhooXhh3kMQ1EBZyasz+UzayeG0uk49EvhGdy9nZf/wX+fts5NnRi0gDYopYlyyzjmUDMcTD8MivbWFQrD+P6ICA6PBqw4tMDCZYt49gXPDnvkbDCgHyEU6uJ2Q6KIsGrVKl784hcjXvDiMbpVycdai1a59kmORnBPW3+RnBhawTKCwfoMowGfkIwOTmubU7874yHyrdQYqnXBmTjCcmJATLCgiSozd7QOjQkSNFFUQemIuNrDyaeeyilnnIbSatyqWbhuHwtov5eCTTLGsGjRIl7wghegdB7qqVq/j+M4GNrU0U0IUAjQUmoHCTuDQKwNHgtZyu6nn5rWdjtIIpaHjCl/gHCzowhFizLa+yKTHSq35ms8wvIVK3nFK16xT2gkULpvzwaXhE5Q7AztNoZarcY555zDmtVrZl4LCOOszYWhrb1jWgEiGG8Z27Obh++6c1qb7yBvUi6eSpFjB/Yz847K4ciNZ4TIO5ksJrRc9Qq3L4USxcoVK3nZpS8fNyHGBcXDUdW3H0kUrFJPTw9r167lda97XZ4xZKZ7ti8KFXnxHDUQoah4YXTXLn75bz49rb3uMInY+AkZ+iwzdrTbLvfL/vvCoiqIc8RKM6/aw2knn8ypp55a6uMLPfyxgonE3Wg0WLx4MS984Qtn3b2224ba/9cEe6qMNti5YXpZpOL6U0QIH5PCpZbCgUrPyKFRaIQIj8mPicQKnjIjgPgw4N6jnOe4lat44xvfEO4s99s5lrLrTSTyBQsWcPbZZ7NixYqj5mpxKGht6EWIVi43QK4nF5r9A9x9003T3nZn2TEIalWYHSxnhJogMhSDGWQbKYlBAgOaq4XxsGL5Sl76kpeWWeXaXRbmSh3oQ0ERy7BmzRre/va3B+I4ypqiA6KNM9LtMkOb8W14527e88l/m3YK7mAECkLIDTXFxqBkRo6ACTuT0uXuZZUi05oUTao0GRqnVLCNmAhdqXDqaafx/Oc/v9StF8m2jhV5oQx9jWOWL1/Oi170orIS6KyK15DwXynytW3sDI3y8N33HpFmO1wOZoMaqQ2TzNlCoij66nKltZfcUTjXYFgnrFy5issuu6zcEY6V3aBAcV9LlizhNa95DfPmzSvTSc6qoiW5faHc54vHp2Fk51Ze/tu/e0Qm3dSJIW3mKy94J6U6WM3QP6Cd4Rxn0dcEX6QYRQ1NDFSUIip+qiGKFMuXL+NVr3oVxx13HBDYoziO55xatYxByCd5kWIewj2tXr2ayy+/vIzoc85NC5/bLp1N6WxxOJ+R4sNC5Qy4PG2Pb0Cyl3WP39V5R/eDqRNDbMrsBEqNS0U0KzGRTibb0wqL7Bve8IaSNbLWzh5++hAwUc5pLzwCMH/+fC655BLOP/98lFKkaRqMcTP4/EqNkXhUUS3J5XKd0qF0gPH0b9nABW951xHraWdP+RjhpQsoFMcddxxXXHEFCxcuHEcQcwEFIUxmJCw+O+WUU3j9619fEnlRjGSmWMKWKlxhUWhliL0iQpGJkNgUEQdjDca2TE9E2/4wdWJoNMrtYK6xEZNCwPnANpx44om86lWvIooitNbUarU5c4+FwF+wSe2EUavVOPfcc8t7m24iONwRknHvFV6ZkIzOerRWeCVEkUZ5T2Prdo5/xRuO6EOYMjFke/YGZjuH93N8l1BgtCGKIk444QTe+ta3lixGEUA/V9Ae61z8XewKb3vb28YJy1oHuU9Pk8fxVFUqwdCmEK/AaKwLcQtGeRgaYsOd90xL/w7WhylhZHgYkJIIjgn1Yx7qaLThWec8i9e85jWlenVWqR4nwUTWqJAZ2v2tTj75ZN7w+je0CFuB+Nlzb8pDFCtEgzfBTRvnGNu4iWe981eP+NbcmTtG4dpwlL0ajwikrRJmZDjhxBN497vfjfe+9GCdayh2hDiOOeuss3j3u98dtH+5vcFZt4+36ozBAza8NFUg7hoCe/dw3zXXHZUuTHkGR1GEOFeO4RycK+OhWtZZ5xzz5s3jwgsv5K1vfeucqM8wsf5cAaUUtVqN008/nSuuuCIY2bwQmQil1YymwhnnQ6nARCHmSgDvElRzjK0PPsSL/vcfHZXZNWViEBGUMSit5r68AOOEyWLyn3DCCfz6r/86fX19ZTxA8aq1LlMxTnSCm0m0124rWKe1a9fygQ98YNIYhaMZs9FOsL4INqK0sYU+44jIqHrPyBMbWPvKy4/aMtsxbzPXN4QSuRBZ5CJ1zlGpVDjl5FN44xvfWH4WRVFZNjbLspJAZgPfXUy2wiFPROjr6+PlL385r7rsVTPdvTIFfpZl6CIM1ToQwVmHKCEyiop42LWLm7/17aPavznO6E8z1HgNTKFm/ZVf+RWWLFkCUPryRFFEHMdYa2dFlmoIxczby9dWq1XOOOMM3vve9+I6LBg+HdAqt2vQkmdahVEcXlKoj8Defp648SYu//MPH9VB7ZwYjgEOqYS0CZ1RXD6s888/n9/7vd8rf1a4dDvnqNVqs8aZzzlXGtOKlJGXX3455z/n/FnBxhXpZ4wxeOfI0pSs0URsiMAzWgDHtltv44xfPvLao4mYlp1htpR/7hh5BFzuy1eyHPPnz+f1r389F1xwAVEU4ZwjTVOcczSbzVmxKxQoiLKnp4dzzjmHX//1Xz/6CYPZNxax6FuWZbnPVEQcxcS1GsoYSFOoj7L1pzez5g1XzMiAdtmkCVBKheS7uYBXZKtevXo1v/mbH8R7P05wnm3x0QUBr1mzhg984AMsX7488OVHsw/7+VxrTRzFeOtwaRp823ywJeAcO++8m7WvftOMDWaXGCZBIUhHJipZoBUrVvCqV13G+973PtI0LbVJs4VFAkrC7evr4znPeQ5ve9vbqFarpQr1aGB/rQSzVG4IdA5TqYBSuDTDZhkbbruTlZe+ekZXlSkTw/gTZ8/KOFUIQbUaioIG70mfl0kyuQZp9erVvOUtb2XtcccBiiRJQKngaTkdRscJvgwH8rCd9EsveOc568yz+KP//UehvoILdRYOaffqkF5a6UWLOPOJ+RhC5R1jdKj0lCWYZp2bvvttTnn5y2Z8Ek35CaqimmBbED7syyvOlQMoJ/TE1+J9FMecf8GFfOADvxnEijxsVCS4NXSEfeIwQgxG69BocqKbQAgqD+qrmIgVS5bxxstfz4XPuYAQ3apb83PS4KcysKwVTbafwATvIc18vnBAlgnegc1CKgZLhuAQHN65UI9bGFdmL3gtCNgm7NrKD/75k7z8l35pxgkBOmaTWqtAgf3FDcz1A8LOsWjhfF53+eU861nnIBIqnSJyxIVURTCOiZdAEIULhVZInvwscRmnnnE6v/P7vxvyuAPO2VYuqQlTTsY3cEBIbpQ0WuWTXIgjhdYQaYXyQoTK5ROFNganoWmFNNSQDMmdfQajwwzedw+ccJq6/H//2awgBOiAGJxywbNKhcDK1vIyk5m4j9yh8BgtGA0nn3Q87/+NX0MjaBUybbhOXRomrMZSrrHhsDhSm8sqcYxyPuRwc4LKgsPksrWreOUbX0tlQS9Nn2G1oGKDVZ7EZUjbPxifZgdohY9PwpspJUSRwmiPJkUbB6Tgm6CTkGWkERP7GloMiYf+tIHECm880ED5Mdizjdu++B8suuQls4YICkzdHUNJIIYOAv3mHoLVdNGCBVzy3OexZs1qvBdiE/K6dnz1gqXIIdAqYaYIDjwIaTNBeYgxxCoq1cFrjz+OX/yld1CJK0TGIHisWLw4YqOLK+aTf/wxcRkb/1Tzv7xDnA3vxeO9zecBgU/TClxY/GMNi6s1YlIq6TAmG+Ppn90Ma05Wl/z2H8w6QoBOdgYTnpxHWivJDOZNOuKHV0iWZ1dWMWvXHM/b3vx2tCjE645zFhcywn7juRVlERGDJlJBihAPsaqwculyXnrJCzjjhJOI8OAylHNUlKaidM4pTU4IgUB8eei23bBIxyaAE0FMhJiYzCusjnG6SiKKhnM4DVRDBh6TWuKkSbR7N3tuux16lqkTXnlkg3M6xdRlBt96SvJM0NDmbpXeCUmasmrNcbzujW+k1jePvOxHZ5cn17a0X6kggvwD8SExWmQijDJ5Jc4wxZcsW8FrX/1axHlc5ohMRKR1PouLhGqTUFi57RQE0doTxnNuCm0ilDJ4NNrEGBXjvSBOUYtjDAKZhTSB/j08+YOrYPlateIlM6syPVRMeRYb0SC5riMfMVHgjuVDg1WCrlZoesuq49byope+BBVHwYp6BBCmaYBWYeo7Z8m8Reso6OvjiNPOOINXvea1oDTKRDSbKSgTprgHlBnPc3nVnuwKJOQkVG2EIigcGofKc9kKVnzIbk5IxF5xihoxKk3AjsDup7nu//0TrFqlTr/iLXOCCApMnRh8CEdSokOSumcARDzGaJSCOI5Yu3YNV/zCz+OyBO87E6AnajZbbbY+KNLiaG3wCJk4mmnKwqVLeO4LLgmlkUVhtKFW7cFZD16hTYy3E3YGpdhnh/C6JIqCGIreaYQYIRaHdmlwnxgbA2th+3bu+NpXoGeB4rhT1Ct/9w/nFBEUmLIju1ifRyYJOi/HqJCQ4/QYhVEK8EgeEbdwXi/POedsli5ZRP/ewUnVCBNDMQ+EVkU1JlBE2BVMLkR7H1KhYRQozeoT1vL2d7wjX5R0SEWu891bq5BiVmtEWiyXzx0Snc8JPBPIBCrBjqEEdFT0yqLFg0sDG+Q89A/x8PU38Kz3vE9x/Gqe9yvvPvwBnWWYugXaRGA0uhLErJY7/yR86bF0SFihxQWiWLl8Ba982SuI98MmHaobxGS7QmGwUoRX8XnmCw09PVUQR6USccYZp3LiCcflMgVtmT/DzPdKcluD4PE4cThxoECbPMugBiINqc3tARaVJaikgU4aMLCXrbffzvf+6ZMwb5nihFPVs97zvjm5A+wPU94ZmmT0GYdVMRKFsXTumBqbfaBK1iKUX8XDspVrePllr+Hr//Wd/Z53aPEO0pKcpWCJAnT+YTuhpGmTSGtWLl3EK178AiLtEaPJbIr3ghKhElfyvjoy5zBKYXSEUkFYtq5ObjMMVrHMQlyBRoOxp57iqcceZd3DD/Pzf/kRRXUha1+0lrUvuuzwBm0OYcrEYGKFSAIKjMQgGqNbk+WYRNutaR2CUkxvL6edeip9vT2Mjo3u99SD7hDF0EnrTz1uLFWwNiMIHuVAi2PlwgWcf+ppGO9AeyoRLQHENiFLiOM4z36oIR0D71C1HnSWkO7Zy2OPPcqWp9azdcMT/Prf/bOi0kff+ct49vnP5dlTGKa5iqkTQ1ZHS4bOLJheIM/vcazD5zKDCBiDVooT16zkpS9+IVddc23n1x9HEKrNGBYMXRBCI6tKwMFx8+dzyfMvgZHRYOkyCvJQSoaH6d/0NOnIGLu2bmXn1m0M7trNW//pU4HKKvOo9C3lvBPP4LzOez7nMWVi2PzggyxOG1iniCUitRLSGxyjO4MnJFcuEvkW8Q5ZlrFz506Omz+PiFBCa8rICUFN+DDIxRoVG/AW7yxGwYIIGBmBeQsmH/T5y1iy9hQAVnXSr2cI1GzxxT8W0KOUJFNNK9PGJqncyhwKMAZTWFkbxqUoYJ6HhcBxlZhbk/TYXIGOMp4BpuOjB0OukmwzXU0FGsrifq2Kl7nbRBxTKId6oEsI04guMUwj5tdqGCBSLT8jo4MYvI82aaLWFspdIXwdCrhLUV6YECKpnAs1rQkiQhfTh+5wTiOU0milMXlQkAYQypjpw0HLPS4/lEKZCJy0iCB+BigsjiKmLDN89yN/Oc5J4Bnjxb0f9KeeG+5/gG9eeWXIpaRARxFZloFWB0/9Ps7G0FqlguucLn+jJFQ0XTmvj59/9Su47EUvIkqaKDS2ELbnIpRHlAcp6nNrlG93WxR8UZyydIaDN/6vP5k2NnHK2qQ3/cH/DB06RrVHh40kZd6VV/Llb30z/C0qtxar0vXh4AuPCkbuPFBqotwhEuwMmIjlxx3HC1/xKt74a++GLAEVQVRlzj4PVSiRofSVamdcyrFrM8RMM6aeZFNNKH1U9P8ZChXFrD7uOOK4EjJci+CtRRtd+jId5ArBcSgnBCHUtTYUtudQmVQbAxpcXOGMCy4AZSCqQFzNd5C5+hBawWItB8HcCi+0Yu6ltUtON6ZODJEe16E89vsZCwVUemqsPf54ntr4VMi/lNdUO7yrFGithIUdToknsx60otLXx/yly/FRTBHmLHruyhCFxV0gr8ja9p0q3K3Gaxymmy+ZMjFYsSAhyq3VqWcuOShR9Pb1ctLJJ7N+w3qUDqu4cy4k2TVmn9TvE1mgdv64EJ1b4xpioZXWqCim1tPHyaefjhVBxzUEj0h65G/0CEHnooAgaC0h1KJgmyRkIFHFZMt3UUGhTDxtfZi6O4Y2qAnOY3NYfOscStHb28uatWtKIijyijrn8AcghPFoaZDaQhmA3MNUhDgyZXle5zyZ99QqMZGKmLMLklIh9iKPrTeBLMqvAgvZvmCoaa8wO3U2KQtO9mItOo7H68ufoVBolNJ478ucrAVRiMhBhqfle1QQgg8XbYtxCGGecRyzcMECNFCLDJBnsp6mumwzguKGQxmocQK15IMgLq8QJSEeQ0fTO+GmPHrKaFAKHcet8NpnOBYvWcyFF14wrhqOVorImIMOT1hLxoW1tWXFoBxfpaBWrbB4wYIQzONbMQ9zGjLxfRERboCQJU1FuhwTbRTeTe8u2EHJFo/1HhVFpULs6NR/mcVQiiiO0VqXdoWCTZoME1nM4IbR9nkbAbRrZXtrVZYvXZIvnj5fSfe95lzCxCjUsNPlwUn5n17KeCW0gNgMzOEbNPeHqc9fbzGAF4co08bJPTMhQJIkjIyMlMRQeLZ26gzZrmIXgTiKWbJwEXleS9q3jjk7/hPMCgFS3tU+xODdtPtPTJkYHr3+es5+6UswSgWWiYPxxMc4BGo9NdauXVvuBIVtIYqicZqkyUgjxJ5N+HYC+6mVCaGdnlABp1xJS4Z7Ou/oKKNdVT8+jVnhwVv8RIuAS8n27iRec9J09mBquPOGHwdSLbfnuWzw6RxKKebPm8eaNatRSpXp6oFDd8Vo/6Pd4OpBeYXJXRW0FNZsBZFGlITMdnMaktuqWuFM7bxh6Z7i89xO3jO4ffu09mDKO0NNR8H6mTqaFUOsg6fmM5kgAlRZcVPl6eq9c3nU2iQTto2/bD37wtpaLP4hoxGiMBjwkDmP6GDs9Hg0hmgO+106intRRLTivxE9TkumRAe1ktJUp5n+p54dw+YFe+MKkdY4P+mjfsZAhLLSj3OOJEnaqoCqkBrysFBMiPFWaYVCGZVn02NSi+3ch24dbUK09yH5OCrYo72b3rREHRjdgCgCHZKIGa2e0XuCUmDyXcAYUwrQhTDtZT+aHtnnzYSvwv+u/N/jdUhiJt6HdPBEOV89dxGkziJAajyHIfn3lcL8UHwVT6/+cup2htzHPlzkGZFt9YAQCVVzCkIAyljp4vspX5ucl85dFHye8wjC/hGJPiYWIl0oYdp8Udpt8RNFKz/Nd/1Mn8PTBqUUSZqwZcuW8u9Cg1TIEJ1eP1TJDHaLer1efifIUS9ieCyiSwzTiCRJ2LFjB0qFKjsFOiUEoKzi6ZyjWq2ydOnSceramShve6yhSwzTCGMMtVoN7z3WWowxJWEcVL16GG3U63U2bdpUZsxTHEIkXRcHRZcYphFpmjIwMABQppT03oeqoNOAwgu2Xq+zYcOGNg3s7Cm/O5fRJYZpxOjoKOvWrSv/drlmaTpkBggGpyzLiKKIWq1GlmWlitUcofoQzyR0iWG6IJBlGY1GY5zRLaSCl2lhY7TWaK3Zs2cPd999N3Ech9K2cJgRdV1Mhi4xTAEikxgNFNTH6uXOULBIE6PbOm23YIecc4wMj+RNq0Mret7FAdElhsNBEWqgVMuFqI0w6o06Tz/99DiWZbp4eaVUSWAiwvDwME8++WRpf5gONuyZju4IHgYm7gjFSl0cSTNh7969WGux1qK1plqtlnJDp20XhFXYGrbvCI5q1toumzQNeMbH43SKgkUZGh7igQcfKOWEYgUvfJQ63SEKTZL3Hu89Q0NDbNq0CWBaiK2L7s5wWJiMPSo+Gx4a5tFHHy0nbDsBTAerVISSFkQxOjLKY489VhJdF52jSwyHg3ZCYDxxFPJCEdlWOOlVq9Vpc8coCM1ay9DwEI888kjpKdtF5+iO4mGiJARUafRyzrFnzx7uuecevPelCtR7T5qm4/j9TtotdoFCmB4ZGWHz5s3jEhB0MXVMWWZQkoFkIKFG8D65Meci9pmvss+fqnCn1EX6Cg9pRmN4mF07d5ZE0O6GMVWZoYiLL5tXtLxinWNkYIAH7rqLU48/nshEoUjhHIVCUCqfRz6P7/Ths/A5rXQyAjhLLNN7v1MXoO0IyChkCnSeu8fPYXm8DLMcpy5qfaFUKPEZmTySx4f3mcONDPPAHXeM2wE6Wa2LoJ5Wfj2KJHI4JWglxEYxunsXV7z97YpmIvg0jwCbUpOzACFMadyCqoRWxu02TV7qwFoqjca09qBbxmqa8OzlS+TxPQPTVtOtIAUp8hSWEyP8JhboBc46/nhue/rpOUsCswldYpgmnKCU7ACyqV6gbTqHpGD7EkMRCqy1JnKOioeF1SonLF/BPJdyzbYdXaLoAHOcyZ89aGOopgVF3lpVRHgpBToKHJp1WB8qi1ql6Fu8uEsI04COdobGA3dIz4knQ7U3CJJmDtOWEOSAIlOV1q2MVZKnrpA81LWos1yJGdi2ncUnn6KqSklGBwQxYWcoRlJychAArdBG450lUqCcsLi3h51jdUWaCDJnYxrqoDYheicwBmwGdqLUGGUFEwHvfgejTwTyTIIqvGaOvXf+jKUvf01HC0JHEm8/lrULewFF5gxmGlP9HW0oUSEQXWkQQVxIua+imDKPS6HeqYbJaa3j0R27+bXXvUmsiRE3ZSYp7wRFhhg8LYIoNFiiVJnNW5RCa2Go0eCVl1wi3/jWf7F41ZrO2p859II6G+HscZ+q8e/bQqPLakjeO6KawS9Z0nEnOiKGyFQhdVgdEUUV5jTXpUK6d3G5lTfO8/4LiPfBIc7nD0OCEc15z5YtW3jssUeZ9vzoRbfKvUZhlMIpTRQbnLVYoGIMW3bt5NobbuTtv/zLR6QPRw8H21eltOOICvmiLBaHUJnX03HrHRGDOAFTI0LhbEgRPpfFcaUUakKQjEAoEEJYjYxWeC80myk7duzgpptuCqzTEQ67FHJVbZ5oIJTGgtQ7tu3ezQ033cRbf/GXjmgfjijG1XTLP4KSPoqK2JE2oSCwd2htMMaQZhmmUu24C53tDKJBDCgzp8WFAvv4uknbTi2AL6zOCnHCY48+zrU/vDZPs9nhzjDJKuIn+4G0RINC9T4yNsbtd97Jgw88xLnnnkvhnSFtfMWs9+NThdF2wkC0qZshvNdKUWkr2RXHEfQt6LgLHU1hySttwfiXuXrsg/YvNegoTLCdO3fw0Y/9I5df/mrWb3iSKNJMt/B6wB12ki/vu+8+Xvva1/DRj/7DOB8qZYqNS47uYE5h8IvsSNL+hbS9jntfnhB+OX8hm390ZUeMSUc7g1OE6j1lAca5zCQVG3ELqu1/UHhxZDZj49Pr+fg//SNoQWmHE0dcMWRph+4B+3p/HPi34wRMYceubXzxy//Oy17xEi688EJECdYJ2mgcDj2LPfYFhSvKVuWfaYodrSU37edkUBpV7YxV6py5UUXSwwJyzBzF/8WrUorBoUF+cPVV7Nndj0iuidWQzUQSLxn/XoDde3Zz3fXXkdmg6DXGoFBEJpoFI3rgY/84ND7P9PUe9DcHQkfEoHKvKdX2ydw9NGq/R/FPMzo8xronNmC0Io40WoGzHOxpHsJYTsKuHQov1zaTtFLs2rmHn/30VipRDa0iQOGsx1l/gPub+UOjyT3cyuJV+9zugehBG3oXLuJTv/auKT+J6Uj1Vt5AwExP6s4IYvL34fAetm/fyU9/egvOCVnmcYVv3DRwiEXu6XHYHxFM/FhAYYijKls2b+Ohhx5BKY14hVKGqLCXzNqD0v3qgKUmDrAwSBQxf9GiA5x8YHREDLptBhyIaOcWJh9pERgZGeORRx5j69btxFEVY4pKAoqo4/SOYe8phce2wKH99XA8uYaCf41Gwu5de7nuuh/n2iSFUpq0U3nmSGO/BCAQlKltR868TljH+ubPZ1VeEngq6GxnmCxlyjEKpWDHjh18/3vfwxhDZrM863aUl6nqfLKN00cdnJGe2MOQqwnYtWsn115zDUkzC0YqBXE8F5OM7W8AWtJcO2GYWg9LV6yYcmsdygyhY8dCoZL2+INWTALla7OZcP/993LVD76P90FIUBqcs2SZ7Tz7RfuDLY79SJnF2/GVzwTBgRIym/LYY4/w5a98kSgO+7dzDu8FkeBGIgJZZsfd44xiUtmoteyrff6N/wwUunceK487fspd6IgYjM4DXbRizrqI5VBKYbQhy7IwsQW8d/kkcmzYsJ5/+ZfPAlKq+1phmC0C6gSTbgYH2CHGK5OKwCLBGM3mLZv5xje+zuDgIN45osiglQpFVYxGqdxYRbEIzDQ1HAyTUsv475VBOigM36EALeD9XPZIKiFesM4SRRHOOtI0xZhg7reZ5aGHHuLGG28sy9kWO8F0pIGZDhRx0cV7gMcee4zvfOc7mChPJOBCfiVFcOlIkxCf3Z4+f25D6Js3b8pndzSPk0ZCof+dBfOhIyitiHPnPGMMcSUOPkAiPPbYY3zqU5/a95w8Y0XxfibRnsysSGK2ZcsWvvrVr9Lf3x8CgnKBXwgEEMURNrPTlgt2ZhF2jE4Wps60SVGrAN2xAOdcmBy5gS0yEYODg9x0803cdNNN5YQpiAYoU8LMht0BKPtSvD7yyCN89rOfDX3POYwim3dBzO1EPXcR5KF5ixZO+QodEcO83nklTzvnaULC6h5X4sD65GzT448/zic+8Qmq1eq4pL/teYxmU96ioi8FO7dlyxauvPJKNm3aVBKBMYY4jlGEe3DWHQOpZvJctDO1M4wMDUGeDHf2TIcpQrUmUpHCccuWLfzzP3+WTZs2kSTJuO+hxRrNBhZDa02lUhlf2iovsPjQQw/xV3/1V+XvkiQJxOIdXjwmMseE3CAoeubN5+a//j9TooiO5vCTTz4JSmG0nvMyA5DrA8JKX6/X+eEPf8gXv/ifOOfG5UBqz2zXXtFzJlFUCCryNRUZu5VS1Ot1brzxRr785S8Tx3HI8kerdkRx33MbuYpVKXR1ahGXnckMOoRIFgbTYwFaaZx1PPLoI3z84x8vJ3qx4hZFBqFlm5gNxFCg6FOh9i2E4w0bNvDRj36UBx98EO88XsL3c19WKJArAbQmmgliKIxus15FfSgQQkkoBXv37uX73/8+TzzxxLiUjnMdGzZs4F8/969oo0uVsEiQjWaT3DM1KLQOgWaJm5r807mdYdzrHEabzLBx00a+/OUvl0RwLKyeBet32+23sXnz5pYmiempNzfTCDuig/nzOfWsM6d0jWnxWoWDeBrOIYyOjHLffffx1FNPBY1LW7WcuYxCxtm5cydXXnnluPUrsLsz17fpQMh0bsBZdGUG2KRjDd57du3axW233YZSqkwvfyywSBBWz23btnHDDTcET6Zc9vEuz/4xx6EVoDX1dGqlhrvE0AatNQMDA/zsZz8btxMUWpm5jHYt2LZt29i9e3fwLVOgj4FsDkXUIVFMpVqb0jXm/ihMI6y1bN+xnSeeeAIIxNHb2zttdZxnEu0uF3v27OGWW25FkH1sJnMVZfiHtVPe5abpCc/tgYTAMg8ND/PkunVobUoVZaPZpFKtHp0CggdzzOzw0iB452g2mzz55BPBT0kENxPx29OMUCXAgzEsWDA1l4zOzY6KMJNmCT10wvmOjIzywAMP5qtlKJjhnSfxKdoEo9uk1sW2ooahExN/ky9bBxLClRzaGE78TZtCr70edOHO3dbF8vOBgQEee+zxXINmyi7P6c1BgYo0WIViatq/jnYGXxSSmEWDOPUIXBgZHuaJJx7HRIbevh4QjzGGnp4aPo/6L9y6UcHT1UQGozXeh2TESkEUR+VvxkXmHLATkwziZLvEhMvpPBlxu0/SREIIn4dXY4KKdcOG9ePKbM1pQshRskcytZvpLL2k9kyWFnAuQgFpUmf3zu0glkbd5kE7lmbdhqpVCN6Fz2OjcxYjGHi0DkRgrcVlBR8O2uTGrYMNUb6yHwhGFW4v4aG3nNLC1qy1aYuUK2IbWrtEEDIlj+ce4aGHHuK8884LLIbzmGNAkIapr/CdsUlKaGVNOgKM7tFEPkGeWr+ByBicDunpjdZY78etxEGe8GgNPdUK3juSzJGlNgTp6zA0TgAnaBU+swcgCF104gDw+9Qwy4MfVZ6hWyRY0fPrtKzMbbeZu3c759i9ezfeC95LyAo4x9HpHXQoM/hwiISnPdPKqU4EBudoNtK8GIjPTfvBqqmVJoqiEAOAQWsFziHekyQpQB6Vm7/L+xGpPGeCeLwcmJMVBH+AtUS1WchLta9IyBxesk26tJZ7cUhbSkmlAnFU4gpJkjI0NMTWrdty4gZr/ZwmiEJ07QSdEYNAmf8yFxJnynSjiv+m2AGH4tzznsN119/I8MgwTzz+BGmWlpFgzjmaSZN1T65jz949VOIKJjKkSYoXT7VaRSvN1m1b2bJ5C0PDQ0Bw/JNCa3OA9gP3s3/5S4SQGj+/P601CxYuYMXKlSxftozenj7SNCNNU+JKzMqVKzn++ONZtGghgmAzSxxHLJi/gHPOeRYLFizgtFNPI8sscRQhU+SzZx06mIAdEYMSRUi02mKRZmpIS8llqh2INKuOW8Oq49bgvfDyV6dU4jiPETbBoENYlQPhtUIMTV6bwU/yIAoiPVi3ilE8lIdZGJgK67ieJOAwlxImPBUJskHufhFivg3WWqJ47sczdIrOBOiSNZr57VXlvehUlPcS4hl6auOT2Ja1SNpnXdv7QnDVbdkZCu1Gq1jh/tFidfbzPSA+l0c0OcsVUj2HXSMI8SJFuhm1jyHN5/KCUqoMbVWE+IfZpB6fKlqjPLXtoTNiQIf6ZrHK6aElvB0JTLQsSu5AFx6wRglT1DC3YAA6KbY95XPbasi1X678D1QUhPIszYjiCF1M9nznKewcWue5krwrbQ/SRqgCZRDQsQQRj8rSKWen6CwlPRIKgwPiBa9cbnuYfmhUzn+3VlulTZ7rofis9f9cg4jHuTxQKp/dZfqXXD3qMkcUx5hKSEhQBOgEva9gVNASeSjzQEGu5vBypCptzQo450izhJ5KZco2k46IwUSBMRGXIXEEOj5iU7FQ4gohdjdoR3RJAApQkzHtcwRhpW77IOd/RFrBRUXSL5FclS0erYJ6tSSIKFRUCdkwgmV8fNWDyTD3eSRjDD2mBmnC8NAwU8me1BExaJeBcigTocQh/sjZGoJTWajrUtGmteWLR3yYHHPZmU68IJlFaU1Zh0qplpdGzgp7HzJzKB2qkzrrEC0hW54IeIfk4+NFyDKbE8XBGMi5TRAiHmszwp45A2xSnNQha4Tifl6hpug6eyhQqNY271w4vEe5UOgvsGtzlxgCa5Q/jtLgkNtv8lJn3gWfKaWCNkhEYUwc7B4C4MHk4rqX4KoRh0Lq3rdZpY/yvR0NKKWJ4wqkzWBfmQI6Iobnveu31VUf/p8SqxpeR2SU9eyPGFpJgQPrUK1W6e3tZf6ihSxdvXrOhqg4Aa80CxYsYsmSJUQLFkCtWvhTgITJXbBHOI8p1KHO5WomAyYv85Fbyn3mUMZgzLFIAu0IHrnauilzCGquhzN2MR7XfudKeflllxFVKmQ2I6oEgnLOEe03LX0xB+Y6wQh+z270shVTupEuMRxDaGzbJtXFS1sxwKUxUHLt2/7OnNvyAoR7dNYSDw3CsuVTupm5y2R3AcDGG38gJP1CNiI9q1eha1UwKhw5V6X1gQgBZgshtHmnH/65SiFRTLNen3L7x5bV5RjHwO0/lQXLlmOWL4OeHvCeky75OYjilmn6GYiSyVMgHeSM7ZwYBncLxFCrceTduCVXO7aEyhLKg/azZZE7fEhbXEjp70XQLAmgNIsvuCR8rvIi7LEBCWV34+rcz+00FYzbRTxkaZOeKV6rc2KII4gqYOKgzTjCtFCoGfFBA6PbFkR/QL/QWQ4FoHODWfDEU+SGzUKV6gQVqWB9zDNbhPXA4CyYOb7PH67TcftvC8V6c3SMBVNsv/Phi3I1nstQVXPEVZveh/XTmCAfesDbYJWtdFxxc+bgJcxxpRTK5HO9YKCLkJFCPaokeM364I8UV44d9qiTtVQhZI2ZlBmsg6pCVStHdF2Wdpdl1RK0NKAjBZiZC6aYBuR30PJcbY9Oy4Nz0ME50SufW6F1CCslxD3rZ7I+xHtIUhojQ1O+ROfEUO1B8njgEMZyBPmknAqstWHShPCv0n16LsenKAHti1mvKP0wxKNUEZ0AmLAAODyCxSJ4HLGqUHmGE0M2OMDWjU9x2hQvMQ1cpgJlEMkdxo4Qgs9+sMKWrsfFblHGNc7dxASFyDBeuSgUMaQtclDYcg/O4xJyx/W5by04PIyTMRS4xiibnnh86tfr2OiWJUJUQdQR3RMOAwe+n/7+fu67776y+IjzIR2M0ZrVq1ezeMlibr75Zu6/736UUljniIyhVqtx/vnn84pXvpKnnnqKL3/5SxhjSNOMaqXCWL3OmjWrefOb38zKlatwziJeiOOYYjp77zF5gjLnPcPDw/z4+ut59NFHsc4hWmPyRMerV6/m9a9/PWvXrGnrvQqOiQhK6Tx6RHhq41N8+YtfAudREhaLNEupVqutVDDk6SXz9DciQk9PD+eccw7Lli7j7HPOZv68+WVwk/iQXMCUgvrM55wtKipNLB+mlII0of/++1jy3OdPvZPtVSKndCRNEefFe5n1yDIr1177I6lWa6K0FqWUmCgSrY0sXrJEfvf3fk82Pf20/MZvvD8sxEqJMZEYY2TBgoXyW7/92yIicuONN4rWJhzGSLUWrnfRRRfLnXfeJd578d6LtVbG6nVx3kuapmU/nPPinJeNGzeGtpQSZbSgEWXC+/MuvEDuuvtuSbNMkjQV57x4JyJexDuRLHWSJk68F7n11jtl9eq1olS4J6WURFEkgGitpVqtitZatNYhaj3nu4rfGWNkyZIl8ra3vU2uueYaES/inBMRCfeR2Zl4XPuF917Ei0jeN++cSJrIwE03inQwl6chJT1t4Yoyg8e47kx6mMgQVyo474PLNCpPleKDe7gyZFlIKalNCIV0ziEoevv66OnpxXkhSS3e50yL0iRJitYGL0JPby+CwjqP82H1HRurs3HT0zgP1vngfq0VXiBJ0zydZfDJR8BojfFQMRGRMlSiOKSEEagP1Xl63SYirYm1RjmoqZjVy1aWkW8iMq7SUJIk5apaoD3znnOOgYEBvvWtb/Gud72Lv/yrvyy/8661Es8ohLI0gHeFPSmUUFMoGBnj6cef6KiJI6CZnkmVzsFZNfE+T/wluYkwTMxatUJkgqZKqzzIX1SZY6hRr5MmCUYpKnEE+Nx51AEKn6eOaTbqKIS4Tc1791138bKXXcrixYt5yUteyhlnnMHpp5/OypUriCODdw6lBHFQiWN85nAu4+lNG3lq/XpuueVWHnrwIdavW8/g4BDvfc97+fBf/3XL5gJUciJXipIFjKKoLGkFgdgqlQpJkpQTvL1yqXOO/v5+/uVf/oX58xfwu7/7O1hnUXrmY0VEkauY8xvOibiI5ksHB9jwyCOc10Eb00AMxbqrcEfcAr1/lAGO+6HFPTt38t3vfpdbb70VLT6oMnXgwZ0XRgcHuPXmm8BmPHjvPUiWEWmFszaYUnzG3bffyl/++Z+yadMmIlre1UoJXoTd27fyr5/9DLecfTaXXnop5573HIaHhrjvrtuJFQwPDHDVd7/DD00Yp+XLl3PRRRcRK8FEhiRzpEmGUrBz5w7+5u/+hltvuS1vIwT/G2O4/b47oNISmFWPIlO2LcGZRSlFmuY5nfJMHsUuUeSAap/gxaprrWXHjh1885vf4PWvfx2nnXZanjnQYY6iHWfiY1SAMjpXouS1BAWKnF1jg4P8/Ec/1tHkm4bEw0HjIWp6slN02Bn2Rw07duzgs5/9LPfddx+VOAoRYLRig9M05e677+bOO+8kiqJg6/JClKuNh0fGuP32O7jjjjsAiCJNZgNRaa0R59m9ezef+9y/8qxnncPznvd8UIqRkWE2b95MEarsBawTtBL65vWxdOlSjDFkmUNrRaVSIcsy+ubN48wzz+LW225HXJ4/VYETRzNtsmXbFtasWYsgZD6j0azvU0eiyAubZVm4p3w1LYikEEaL1TaKojLb+Lp16/jG17/B//gf/6PMKXu0MNkTFBjvfkPOVjoPYtm9cyeLO2x32ohB5SkmZ84GXOjnJ18cnNJ4bVBxhWaWYaIYh2DbdxMX7sE5QUdRkCdEMCbG+SxfYX3506hSIU1TrMuXAC9UKhVOPv0snvfCF+OcY2isySNPrMcSEopFcUyWZcxfuJCzn30eqRdSD6I0SkU0mhkK6OmZz3HHnYRWMWgXtCe538nQ4BAb1m1g7Zq1iECkdJAr8jQwSqmSPZo3bx4veMELys8HBwe5//77g8apTDoc4h0Kdsp7T5qkrFu/HhPlxHKU9Lb7Z7KlNQYhdI8iYdRgfz/33HUXZ7z2tR21PW0yQzlWMyUyqFY/JoOXINSmaQYQiuEpQOmSf5a2yeAlT7HiBS8Wnf9G5apRlCbNbOtvAaUVS5Yu48KLLgoq2cgwNDzM/Q88AEqjtCHNAlHVenpYsXJVaJfQD++C/KGUorenl6VLl1GJYxqNVr1p7zy7d+3hlp/dws/93M8FeQOFbmOFoMUaXXjhhfz4xz8uBehih/DeMzIywre+9S3+5m/+hvXr15dp9b33DA4NsmHD+nxo1VHlfkNw68QEaG2FJn0e1eZDioihbVt5x5/8Scc9nMZiJW0r80wc+/Zk/KFyg10eKqk0QfsiHudClm2tQ02GwJPmxJFbgb2z5WcKCa+KkKtHUToSLVwwnwvOfw5Ga4YGh3jooQfZsX0bxhRp6z0inkULF3DO2WeTZWkwrPsgRAf5w2FdxuLFCznjzDMCe5TXl9NaMzw8xP0PBDtIZKJcE6bbMuyFx6p1SFWfJilat/KwQhBEFy1axBvf8Ebe9a53lerF4vxxOZUO14OuIwi+yL9VeCbnYqkLWf9xXoILjvGQjDH85CPT0vL0qFaB2W77DLtWWFEqlTj4+PgQL2tyjREIkQmJhU3+qnP2ophYLboLK1ccRYjPU7YoxapVq3jFK14BCGNjozz0wIMAOGuRnEfvqdWoVqusXbsmaKFE8sTCqnzVWjFvXh/nnvvs8jPJCSlNU9avX89TTz2Fy1kHLzJu1S8mt3OOSqVSfpamaak9giAnLF68mKKqafG7LCvYQmkbwKODEMU9XmVedKHZzPCFys8nNLZt4f4bbpyWdqdnZ5gDDnKnnnoqn/3sZ/nIRz5SCpTFJHfOMX/+fK644gq++MUv8upXv7rkoYvJMW/ePN7ylrdw7bXX8o//+I/lZCu0MlprzjjjDP74j/+4LJm7bds2rv7h1eWEMsaglKLRaLB48WIuvvjiIHzn3xdtFhN6wYIFPP/5zy9X7EIgds6xY8cOvve975VWZaDUIk3MlJfZVh8rlUrZXpZlbN6yma9//evjCCTkaIrL33l35Ev/CrnXLq1s5uPkPw/aQa0aI8rjfAZaM7Z3L+/8zL9MC6lOm8ygZrljzLy+eVx00UXUx4KLb/HwC+OUMYbjjz+e5z33eVx55ffKyWOtDaxGmnLTTTfR39/P8PBwqXlp1+VXKhUWLVqEMYYkSdi4cSPr16+nWq2SJEmp01+0aBHnn39+qf0piLIghOJ9X28fp5xySingFv3WWrN7925uueUWPvjBD6KVHqdFcq6VIWLLli38+Z/9edgdxBOZiMxmbN26lYceeojHH3+c4eHhst1iN+jr6+OMM84AgsuHPppRdF6187et8h8msKWJT5mvYfTpLfzs2ut40wtfNi3NzvFwkMOACquOF18antr547L2AUKtVi19X/r6+qjX6yRJwp49e7juuusAylU2TdNyxY+iiEpcwVrL05ue5mtf+1qp228nrgULFnD22WcjIsRxpdyB2q3CSil6e3sxxnDiiSeycePGkoCzLKSe37hxIw8+GNiwdrtBMbGttWzYsIGP/N1HSgIpdpZiIWhf8Yt2vffUajVOOumkQMBHe5XL3Y+lXW7PHTWdciFuxaeMbd/Om/78w9PWuWeUz2+IhdDjCEFrTZZl41ZpoJwsY2Nj5d/FJAryROs6EFiULMtIsxSFYsNTG7j22mupVkM278LaG8cxy5Yt46KLLkJrTaPRGMe+tAuySisWLFzAxRdfXBJsuwywadMmrrnmGhQhf1S7JqnoV9Fu8b4g8izLyt9WKpVS3oAgR5x44om86rJXhXaNPuIVQfOFv/VHMcVbNl1UrIgiTYwn3bmT23/4o2ntwzOKGEQkX/lrJa9eTIBCbVlMnmK1LyZqwY+3C6gFURS/r1QqxFHM9h3b+clPfsLIyEgpsMZxXBLf2rVrOf/880GgtzdE7GZZVl6vaCfLMhYuWMjzn//8UCsuX92zLCOOY/r7+/nOd77Djp07Wvx9PuEL9q/YTYpdoV1jVMgFaZqWLJ+IMH/+fJ773Ody4YUXttw1jobRrfC00Owbm6KgiSfzKTpL2LvuKd74F38zrVvWtDnqzXYpupjECkWz2SxVlcWqGUUR5H49wTU7Lc+z1o5zVqtUKuOMVMC4Cbh161a+8IUvjBN6syzDOcfixYt50YteFPh8HVwm2oXm9tVda83KVSu56KKL6Onpad1DLkNkWca2bdv4j//4j3KHaWe1iglfXLMg4GInLDRGheDsnKOvr48rrvgF/vZv/zZYntss1Ef2AVFOIUsuTBfNBjcxBEcFR33LVq75ylenvQudywxZBnH14L+bYZRCqnjiOC4nbrFDFLLDggULeN3rXseJJ56IQmGdLVmTA00KhWLlqpU0mg0++9l/YefOnWU77SxYrVajVqvxo+t+xAMPPMD111+/Tz8n8vHLli3j5S9/OVdddVXJ5hS8/ZYtW/jOd77DsmXLynPaCbeQJQqil1wF2767AVSrVc4//3w++MEP8ku/+Etoo8tFoNg1jjhys0KqoAL4TFBGkbqMuGKI8dBosPWue3j3pz837YJM53doctfJ4E3WeY+OMJzzpeBYCJxRFDEwMMBXv/pV7rjjDoocroV3JzBuF9gflFJ8/etfZ2RkhOXLl7Nr166SNSom4cDAAH/6p39KPU92VWhu6m3Jrwq2pmDN1q5dy6WXXsoPfvCDcbtDFEXMmzePM844g/Xr149zpyiI6rTTTuOd73wn0CL8NE3LnUkpxSmnnMILXvACTjn5lFDzgcCiGd1SDBxx5GyqMoYY8kpEKoS5GoP1TYxtsPXe+zj9re88MhOtk2AIEUGaoyIuE8mDQWYrvPeSZZnccMMNonUrCMYYU75qrcu/IQTGVCqV8v2BjuJ6559/vtx8883S398vX/nKV+Tyyy8vvz/ppJPk7/7u7+RDH/qQKKWkUqmUQTjtgTZKKbngggvkgfsfKIOEbrv1Nlm1alUZrHPuuefKpz/9Gdm+fbvcd999ctFFF5XnF0E8gLz0pS8V73wZcCRexFor1tpyXGxmxVkXvs8/ax+zo/mMkiSVZuZC8I4VEW+l4esiMiLNzY/Kd/74v4l0Omf3c0yDzJDvCrO8LEyxwilUyR60G5YKGaL4LeQOa7kA3G6hneyQfKdZsmQJF5x/AfPnz+dtb3sb3/jGN7jqqqs477zzWLt2Lb/4jl/kkksuYcGCBeME13ZVZ6EtymzQ+GitOeGEE3jPe97DiSeeyD/90z9x88038/73/0bJHkku2xS7QiH4l7Uc8p2iXTnQbDZbnqvkQTMTx8xEiJdS+3Sk4EXInMVEhqrRuRUuQ2yK8SmMjPDwTTfzpr/6xyM20aYnVYz2nRdTOwrw3pOkSenVCW2yRM5iFBqbglAKbUrxeiAsWLCACy+8kJ7entJIVavVeM1rXsNLX/JSBocGWbVqFWP1MV74whdy9dVXA+Mtx865kleP4zgQi4lYsXIFf/Inf8Lv//7vs2TJkpKNKoT5iURbGAoLiATDmVOuJIhaLa+nIWC0aZXFajvHWksURcRxPG3PYTIopdBRDN7jkhQTx6BDAZbYOZ6++x4u/KVfP6IrbufEcDQ0DdMErXWp9y+0JHEcjzOKSc6TF4JnuzpSDnKvK1eu5NJLLy1lkUJNq7Wm1lNjde9qAHpqPRx33PHjnOrafysipaCtdauOXU9PD7VaLfDWk7hPQEtj5L2nt7e3NDAWRGPaavChWrKQMSYoDHLCLHaPOI6PmqLQ49HiieIYvAMcNJtsuf9uTrj0dUec9ZiGjHp5gYxZLkCLb7ERBUvjnCsNZ+3ObcUq3R4yCRx0Z1i+fDkve9nLAnEpTRSPX1GLNuYvmM8555wNtHaigvDiOOb5z38+73rXuzj11FPLc3WeYrLU+efv0zRleGiY++67bx+Xjnq9TpqGLBntRKJ1EJ7bWaqgA5Fx/S0Ipnh/pEM/NZrI6BABpQTEM/jURo578WVHZWJ1niomGRWiGh6FNrOfVxocHOSee+7NPUP1BCEKent6GKuPlfKFy2uoaaXzeOfJEUcxq1av4rTTWimsRFrhkoVLg5cgY+zYvoMNT23AO8+8+fNYvWo1q1av2ofVaVdrirSCbEqhT2mssyVbV0xway2V3NVDKVWeU1y/TAWTP7Pgiq6DA3WR0oYQM74/IpjoU3rgGdtmNNgPvPggxFoHWcrI0xuYf9a5R22F7ZwYmiNCpQ/JA+u7mJsQWiWEYbxHxGTwtEJ8A33qScXGca7Y0k4QqnUhrULkmuQBC82U5IlHqV703KM6oaZh38uHZBazSF0cHCpf2RVhUkz+NFvhUsHVOhxqkl2hNd3Hh1k55/MjzyoiPvg9eQfewugIm3/6k6NOCPAM803q4gAQUF6jvAavcwchTZ7amVYgZk4MotD5EXmFCdmiAIvKK3a3rq3AKWwGWscoFaF0FMJmi3TqWmC4n3u+922Of+3UheXfXmQunuq5XWLoogWZcEz4atI/2ud8yTwVYZut70WDijROQSYeraFiFMpZjLckm57iqn/+DBf+yvs62hH+76C767cXmQumcu40yAxDQmV+nt2hs0t1MYOYSABqwiv70Mc4hPibovixBjEt11MFmYLCbCcuo1eBcg7qDXY/8gjLX/CiaZ09v73InAVE/3fQPXSo53SJoYsD4kCzQ034jZRp5IpSY+GlYJoEMHi0z8BmsGcPN33r27zkt/9gVsycLjF0AQRtkhunTxqP4tGGaZ4LzKL23U3yyZ8zSuXHkThUlgQPvPooO+66i1WXvX5WzZguMXQBgJ9ADO1cks4/Ha9unUAI5P5EevxOEOGDYO0s1Buk656gcsELZuVM6RLDMwiF0a14P86Ah+B1WPNDqhbyGe1zi3CRrSJMc9HBghDEZQHRGCnK1Qi4FB0BNoFmg2TDBqrPmZ1EUKBLDM8wFFbqwj1jXFYNACeI82gBbXTLfhSqKZYF132emtPl20GMJkIF67FNAj81OsiWe+7muFe+cU7MjC4xPENQeOBGJmoZf8eFrYKIJjJtG0CRooX87yLJdJ71zuV15wSIXIppJqA12VMbic+5YM7Nhq6d4RmC0su1WOgllLbSKjgJRkYT6RYheA9ZkQdWA8rlh0fhMd5TySyVZko1yTC793L7ld+GnsVqLhICdHeGZyRKz9Y85sI6Cz44GyISDNBKECV5mQGH8hnGe8hskCEygW3befhnt/GsX+/MUDZb0CWGZwgKWaG9YGHhBh4SC0vQ+IgPce2S8/6Fs1IzIdm1i8fuvofnvPVXjskn3SWGZxLaUoC2ywsignEZiA3OcgADg2x9+GHWXnbkg2pmCzonhrEBodoHSufHhLE7lOuPOyXXUEvxRWHSF8o6zx6CXjtf1doT7LRfZtz1J3um7c4zk/Vpoi59YjPt5xYr7MR2WvbZfbugJvl923mq/fpt19nf7/f7s2JMfcjrDqANOEeyew8b129k+1NPsvWJx/jlv/3oM2byT0TnxDCwXahUoFIDn6siisJ7hUaiaGOyebKPv7ADstxrMsoniwaTQZSE61kNEofvtYQ8/ZBbRCd7lgWRtnxlQv23wqnM5f3VbX1pcz1uv6YqJnXbrCulzDY3hHIC54eSECtenF/47kjhJZpLrZHOkyvkfVMSVmvrglqzCNovfuMlNOsc9O9h66bNYB2RKIxobJqw6rVvesZO8MNBx8Rw45//b2n2VEhFYXSMV4osn1ORBy1C1D5Xy4bDq9P53M4nnsGhJUWJQUlUuhHbKKUZN4md0JMZjI9BYjItZFGeoVpafvWqbZKKUng0XrVCTfJP0LmFVIlGiWrrY3A5U7m7coHwfWvVVhJ6LiikcFbI57vxLWLw2uO04JQPtCUGnRuqjDMU5qo0S0NkXKwYHOhn27YtDA8NgbdEIlSiiP/5Xz/oTu4jgM53hi66OEbQtTN00UWOLjF00UWOLjF00UWOLjF00UWOLjF00UWOLjF00UWOLjF00UWOLjF00UWOjhMPP/iZf5QVp56GqvViMVS0QqUJA09t4LRf++AhWUrv/dwn5ITTTkdsyDEamQhl6+zYuI6z3//fy2sMXHeVWK1RXiFaEJWX0RYD3iPKoXWo+BLy1nrEOazz1EdGeeShh3ndX/zdYVtv7/r038va00/HxzGxjtCZD7lXrcVXNA0sNR3Rv3kH5/zye6bdOnznf3xOnnPhxVROPCV4acRRy6VE5z5RLoW0Cf2DPHr7nZz9jql5lg786DpZ/Jzz8jYEeqrBlcVb0FFw//ACg4MMr9vI+kcf44L3HzhV/N4fXymYHiTViHLYKEOUCzZ+H6N9jFeC0xYnHqMU1biKEsXAzt2se/xx1j/8ML/1pa8fWct7p9VORu++WSTdK9LcK5KOiozuFRnZKY07f3LIFVaGH7hJpLFXpJmINFKRRiYyNCgDt984/hrNMZG0GX6XNEXSMZGkIdLMRNJUxDVFbD30ozki0hgRqY+IjI2K1EdF6mMijTGRjRvk+r//20OvALN3q0hzQCQdEnF1kbQukjVEklERNyKJ6xexAyKbnpDv/cn/nLbKMuuv/o7I8F6RtCGSpCKJF0lFpCkizfx94kSSTCRNQn/S4XD/6ZjI8IDsvfNn8s0P/eEB+3TNX39YZMsWkWZTJLUiiRVJnUjmRTIJRyoiTdd6bWThWY00REYa4rZtleu/+K+Tt9PYFsav0RRJxkTcgIjfK+L3iNhBkaQe+muHRNygiBsSscNhvBuDIvWhcGzfLPf+279M2/hOPDq+wN57bxNxQ5JJQ5xYEd8UscOye+JEPsCx48GfiiQDYlPbKl+UJLLrvlsnXCOTTCT8RkS85KWz8pc0P/Kvw5u2w3sRsTYQ1MiQ7L3+R4fWx2RM0saINCWTTLw458J1vJNMrIxKUzIZE2n2y5brfzA9D2vX0yKje0RcU7zLQpkpm0/MJD8yEW/DnBXvRWwq4jKxNpWx+phkNg1EO7JH5K6bJu3Xlu9/V2RgQCTLRLyXzFppNBJJExvGNQtrjGtIIMI0H+/2w4uIz8JCsX3dJO0MhA5bycesKamMSiKj4iUJ5zsnYutis1HJXF2cZJKJk0ycWJ+FBWFsWGRkUAZuun56xnjC0bHMUFEVoIK1CotBVBVUFaUPpwJoBHENFWnSLA8wUR7lx5dO8iQ4Urw4vFgcGV5sCFN0eSkq51sZDvPkPU4gBZoK6uJpuhSqiiXnn832q791QOesxr33CF4RV/tQhDT1KtP4hgeribyhlwqGKlLppWf1msMcwUmw4RFhfh++Yki1o6k1qdFkJji4EuWHAWU8qAzrMnAKVISYmKinF2sKz94KA/XmPs3c9umPydqXXQo9PVgE61Iio6jWIkxFY5VgI3AV8DWwFXAxpNbSSBq5F67DuSQEBlmBoca+92MVYvOESk4jGDwxjggvJtS6tRp8BaN6iVQPQoRFk6JJVEQa13C1HqjELHrWOey89nvT7lTXMTGkSQJoqlEVBWQerBMOoThmCe/DU9ZKEccRNve9VxO6F/xXVchTq4rPPHiL9pYajopkqKwJSRPSBDKLcUJFgrd3FMXoWg9EFZjfx+ozT+dL/+039juwPWeeASq4khvAOFCNFFVvgvVgBSUK56HhPUtOOpH7v/GVKT+on33m48LqVVCp4OMaFk2GRWOpYNG2CWkDkjGoD0E6TOTqRJIBHjIHqUMLeCugFTIwwOKf27fgxyVvfysYQxbHEMVBNnAZKkvQWUKUJUTNMeJkjChrEGV1TGOYikrpqWiwKXiHUTq0ax0ju3bve1PKhPiJwssdg0HTgwqhpL5wU88z7WUpJkmpOU81PyUDUmXAxDB/ISvPPJPv/cn/N60E0bEArUzIq6+wGAmF8kwcY8yhyzpKVcGFrigNPjKQGbRUxv8w1cSRCT79BpzyRM0mO+97mNtvugGDA5+CeLQoFs5fyLOefQFRdR59p59FdeEiMu9QlZjMO2JdgyXLOeM5F03ar7u+9xW5+PIrQIVqNsYBqaW+ZSNX/+gaXv3mNzFv7fE4D04UkTFQ62H56adPaSwBnvP850HFIEphfai11oPGjw1y780/4YLX/MLkA2vgJx/7WznjggtZfd65MG8+kQYyR3PXTnrWnrDvOX09IdRDa0QgUhE4S3PTZq78/Bd424f/+oAPMXvyUYlXnwCmElb3xjBPr3+UZ/3cy8b/UEd4bCi+ohUOTQ2BvXvZ9NjDkDTpiWOUtWRZRlypUOmdx8KTT8HMW4iJqqBjMq0CMYiGeYs58zkXTnGU94NO+az+u24WsWMiPhXnA/PuvZe9d912yHzdjoceLXlgcSJ1EZGsKf133TX+Gs1EfCPnPcVLKqkke7fJjz/9iUNq68mf3SqSOclSF3hs1xTJRmTzDT+c9Hzf3CWZa4j3It6J2EYqMjAoT3/7OyIiXPWZT4iMjYjzBSttRVwiYzu2yXf+5i+mxNf233mziB+R1Dcky9lxsVY23Xf3YV9vw81XiwzvkD1Xfn3yc7MhsbYuDfGSiYj1qWT9e+Tm//zC4fd9/VMi9901+XlpXbwEgSeVTEYkyACbfnztIbUz+OgjImkm3gfRxjYzkaQpW3501ZTGeH9Hx2ySzvP4CxqXZ+b3XoVIrkOEtY4yjZsKyWkDmzSxbJSgqhW8hsQ5FAYRQ7af0hoTsWPvNtCeyAA+A8nAJoifvKyrj3tJdERTQYrHVAU3sIO7b7wGgMs/8LsKlVBPhlC4PN5H07tkERdd8oJDvv92VEwMTlA+8NjKA9ZwwhnnMfrkk4fFFpz84tco5q9US9/w1skHSOlQHwFFRoZXGdHiKpe89hXc+rUvHh4LcspJiudcNHk7vopyBpd68JoKIWhLpYdWTnfhWWcrRPA21OUzkYFmQt+8+YfVxYOhY2JQZfqEVvyz2k/o8P57odqC1QVVJC1U4wWPsV07Q2JbBc4L4jyVnh5WnXTSQZt44prvy4tffzkoj895XZyFRpPNGzbu8/u9994lStfQRDgBpQQkYWDXVn7+458uH/q222+nr2ZIsybiKMdh6erVhzMCJXZt3gw2I9Ia7TzpWJpHkxrmHX8C1EeF0UGhOSKM9AtDQ8LgmDA0JuzZKzzxmGy6+nvy0099/OCPIE3RaUoVjwEsQiqCWbiQF7zlLZA0hEZDGBoR6nVhdFQYHhDGBoTtG6Vx109l/X99Sb7xh791wLbSXXsAMHEcKv20kjEd+sBICHutVqMwX6oVvJu4WHaGaSiKHpbz9tHQZRjnIXZCE1bqUD4mXE15vBo/xv39exHvyawjjmOMikDHnHfpyyBpCqkVkkRIRoXmgJANCs09QjIgZ7z6FYwpw5ASXKUCcQymghuss/XRDfv0aek5z0aLoQbUgIr32PoITz7ywLjfrbn0cqVQ9MY9aGWCxkRretas5tGvHubqCpz8xrcosgTGhtCRUOmJ83QtHqoR1GKoqBD3Pa8X+irh6K3A/Hlw0imc+PJX8eJf/01IG9J8+gn57Id+b/J+9CxReIvKEowTrI1AzSczfTiJIeoBU4W++RDXoNoDPfPB9MCSlfScdzGnvv4XeOuH/x6G9sjjP7xy0nb27NkNIgg+2AizsPKpWs8hjUn/0xuEnCPIvMdnYa6kjX01ZJ2gY2LweUJaUaGAdchTyD6r+oHQHBsJJXRzhGQTgtfjx9aYUKmyEhnEQ9LMUNVa0LwADS+kUYxUeyDuAVWFynxcVCOjAsRoYhKXJwDIPP1Pb+Gt/zg+I8Rd3/jPQMsOfBNUAqDYuXU7L3jP7+9L5YPNshKHLzKEVKqsPP2MQx6DcZi/Wj3+s59CfSy3LDfAJeGQNLCUoSwm3mdgJCQLiKKgJatUwVRxPqK25gR+7Y/+iLu+9P8mJ4jaIsXWzZikyXwUFe+oeMGoPOGAEsjSVtYTbaBSCYJspQqVCtYYmLeAM1/5akYeuHOfdrzxEBULnccI4IXjznkWT954rWy4/nvSf8uNsuv6H8mOa66R5J47haEBYWRQyJqyZO3xCBovYNM0ZAUfGWXnls1TG9/9oGNtUhHIr/NUGErCq1GHvij29US5bSEKN13wTBOmnVIGl6aYSjXkv63EgOARdBzTo0yePYOwY+XslxEwXqE8mEjjjYakSf2hx1j+0pftM7kvfu0rwsNzIY0icQSZY2holLWT9P/Wa3/GC37+jaHXJuf4lKK6ajVX/sWfyhv/7P8cthvBma/+eQXwtf/+h3LmBRew5pQT6Vu5iL4Vy8PCIR4caBPB6DDoKlR7QTwigAhGh4LmuraAcy55IV/5zd+QX/zMv+zbl5POUAA3/sNfygmnnMaylStYcMZZ5e6JiaDRDOPgVaj7rcDZDKdBRREpGmdh/slnsP1HV8rqy9qSDVc8qR8LbjLeYHRPcO1YtpTTL30ZGQlRAkp6g4rWW4jz+uLWQxzjrISyV9Ua2Ix0106e8+4Du4EcLjovip4zSGH65mlLvCDm0DedFaefGgbAhGz+viAwN/5eV59+alj5iswzmpYWwGXjczSJgNI46zF5oe8IIPGYtMmTP/0Jp7/2tZMP5rzeUJPYW4grQbZAWHP8CTxxy/XSk1oqYjBSxZoq5553EdiQ6aPI9qIqVfpWr+TCS196yOMwGd7+9/9w8AdeCy+3fftrcskrX43qnR+scwpQMSlC74oVnPncCw54mUv/8E8OeXJtvuc2Of688zA6JiFsjCoy0DufvuUrx/32uNNOwmqNYIP62QNOyMQjkSYhplatEGcqZPPTeWF2ASKDeCEyKlhPvYOhIZ689x6edc6zD7W7h4ZO1VH9d94i4jJxElwTMpcFk/uWzYeu9rIjIj4TmwUFXF2cSGNEdt8wwexuGyLeiU+s2MxJ5kVss577IQ2LjOwVqe8VSQfFJwPipSmJT6TuraTOSXNkTO75/uRq1OIYefBO8TIkTTsqYq34RlMks8HdQJq5u0lTJEuCH0+a61UzkXoW/vSSSWqb0vCpDD758NTUf4/eLnt/eojuIm3H9/7t70VkNLgUeZGGExkWL+KHZesN39znej/9h78W2bxJbv6Xzxx+P/v3iDQbYiXct/Oh0eYTD42/lh+RTMYkkyHxviE+cSJWJE2tZGmazx8voxLU6oWW3YsLOu0sDX5N9YbInj3y5Le+Ma0q1WlTrW544jFo1hGbISKgDU5pouWrYMu2g/NK9RHxGEQZlAl1gmPA9g/z+F33TyTdwD2ZsOpZn5E261Cdp6jNV8xbouhZoogXqm3rHkdcAxFb5hOt9vVwwnEHdpeYd9IZQBWNQcSjNEG/6SyhcR0s0iYKW0BEzoeFXV0rUC543oqKWLh2DY9+9T8O31J68iksufgiSIaFvTtk2003HPQa3/rr35NLX3YpqXeBW5SgnDAIzb0DbN20aZ9znvPSl8PK43nxr7wX+vcITz4sT37zSwdt65H/+rpQ6wU04qSV2DBrMjI0MP7HpQuGxjpQcajyE3shQkNq0UmDXt+k6ptUbJ0oGUYldciaQW5qjLL7zttg6VJ12hVvOSLeq9OTN6k5JMQ9IBqvIzKgAqgkDVoPacLYCI8++igL5y9kzenPCluhC9n3vFF4nbNG3qGB5LF1/PuH/pzf+G7LbVeyEZG4L8h1WmHxqP49REtWTj442bCgY1JTQXvQmUOLsPnmn3D8Za/a55w7vvw1ed6b30KGJ67EeOfQYikLMykJ9pOCBdGSGwJcsFLrCuI9Sjypd2RxTJ/LGLj/PhZfeOhVa7bdd4esOedZgeiSDKq1wB5olcsKNiQFjlVQOVoLtVrOVvTgJMa6kG7e25Rqj4btm/npv36OF3/ow+P7Ua+LRL1BnW0bQU6wOcsZV1oZETMXLP9RkNPQCipxXrbK47wj0gbtUwbvvYNFz720bMc26uJ7ehAcOEUsBp1mjG3dwjXf+CrzxWEihTMS3Hucp6cSEynFwK49vOkf/+moJE2bFmLYdvN1sua5l4CJcSom0xqtwWdCpByRtlifoiONJg5OWU4FIUoFT4xgaLOQZjAwzM1f/jo/999+b/wg+IakXqG8CruI9kR79mCWHzfpYN33xf+Q89/8ZqRWQykTUjQ2G9Ac4cdf+RIv/50/HHee3bpdoqUryAR0bNBGUD5j5Kn1/PC738aIJ/KK2GmchkwLKAc4+pYt54KfewXL16zFpQ6zoI9EQRUh27aVeM3Jh/5As6agwKsIraM8TWcujxXpMLUiwwGWGHBOo1QVZxWx0a1CI9YCTbLHHiU+/7nj+jD65AMy7+QzsVLFxHlNhiwjqgSeXVzOwueaNTwt/aN3oAWHR0c6EJNPYXAv91z5HS58d9uzc4jNRQVDng00sWy99WesfflLZ092wGnjufZuF3F1cT6VsSy4WlsRsZmT5uiQiKTipS5Z1hDJnCRjqYgLvOGYWGnaYZFsUGRkt2z5/pWT84TZWHBTzjIRm4iXpsjuLQfmH5/aKpI2xbosd3PORGxDxtY9Iv/+ux8cf27WlDRJRST3KhYvMrBX7vrC5w6JR/3KX/9ZkF/SRCRtSHNsKLh8NEdl108O0e14/VMijUFxdky8ZGJdJt764IJuRdJm8Nm2zkndBzdnL1k+umE8JRHxQyJSuF3v2CE/+/g/jGv/1s/8rYjfJVk2WLq4J06kYQPP3nAidRuOsTyMIfEio82mWO+kmTbFF21ndfHZoEi2W/rv/fG+99m0Yp1I4rLg/u5EZKwh26+eJnf3aTqm94LbN4jUB0LAjfXiMh+ETJuKJKPis+HcOV5EUhE/5sQ2U0ltPQR17F4nD/37AQS54YEgTCVpeE1HRXYegqCejIXAl6wpPstFtGxMhu9o+U/t+dmNITAma4pPrKTNVBrNuow89tDBrz9xUaiPiTTqIW6iMSpiE0kePrTr3Pl/Py726UdE0gGR5lAIJGqE8Sx8t8QGwrDO5QoLK857qWdekoYLBNCQMLM3b5GbP/mxSdtef9dVInaPyOigyMiISBYEW5/aMkbESxFC4SQVL02XylhjWLxLgiKhMRxiJkb3yq47JiEEEaSIjfB5AIbzIvUx2XP192YVMUxvDPSqk9VDX/sm9A9Co4F2ucuDONAeFYUUvdiQVVqhMZkn7h9k3dXXwrJT1bN+5QP73zajOEiFnsC/pna8OnU/2H7rzcEPydVRyobt3GbMP/0Mtlx7rQAsPeccSOogKUossThqCOnOSVySD4AnbvwJpZOV84EVFE1l9Vo2f/HLB+3sxb/1e8ocf7YiXqTu/+pXYMO64CqdjoFrBhV0lqKcxQhEXoH16KRBDwmVqA7REIw+xX1XfQGOW6te/DuTGAqBUy66XGGWqh//w18xdMeNMLgTxgZRtglZA5p1VHOMyDapuAzjmlTx9BqD8rm7eNOSPfI43//Ix1j+3H1tNkCQRVwzCMM+JShiLdFhqN+PBo5Y4uFPvuudsurUUznx1JNYsWo5tb4aicuIMLi6Q5qWk970htnDL04jvv+RD4v1gjGa1AspmoUC9Y2beMtn/3nK9/yNv/9LWbBoKWefcVYIWrKWarWKbTZBLFs3rWfz1o287c87q7Hwjf/+O7J05SrWnnwSC5cto55lVGo1rHiU9zTG6px1+X5cySfBVX//l9Kwllgrms5har3EY3WSbVt422c+P2vmQDcLdxdd5Jhd+1QXXcwgusTQRRc5usTQRRc5usTQRRc5usTQRRc5usTQRRc55hwxfO8v/lhG771NGN4tNIaF+rCQjAm2Ho50NMQHjw2FGOG926X/lhs71h9vuvLbwqZ14ZqNkdBuc0RwDcE3hKwuNOvC6IgwNCjs3Sn1W34yLXrrmz7zT+IfeVDoHxBGRoTGmJCOCX5U8CPhaAtxZc9WefqH3++47Qe/+Flh46PC2F5hrF9ojgppQ7BJ68ia4Ujrwt5tsuGq78xdXf1Mm8AP9bjlU58Q2bstuCjUh0RsEvKrZqnYpCHeZSIubfkuZWnI4ZmN5flRhyRd97B87X/8wWG5AOy965bg9pGMhGul9ZDntV4XSTLxjWae2THLc55mwRUja4a+NgdEtq2Xn37qHw/b9eCRb345d3fI87vaNPg9ZSFViiTN4GuVxwSEWIuGSHNUZGxEZHBIxn52y2G3+9gPvhbcNGRQRIZF3GjehyTEdliXO2/5VkxHPRGpN8O9j+4V2bpevvt3/2dWuVsc7JgTRrex238mfeflmaGjOOTPcIK3Hm1MiIYCvLcoJ8TKkGUZuhqTuBRlFD1aQ5bByAiPfe8qzvrV9x7c8rlnu7BoESkKHVXw4nBJSk+1B6wPkaqVCJeHC5u8tnnmHZFROJsS44M7RSNh1z33sOIV+4mum4iNTwgrV0FcIROFimMsgkGBdcQ6ylNoejAGjyfzITVnJQppMLEekgT27IaTD9FrdtsmYcViUELDGCAiyjPgiVdorbG5T5sWiNqLwAt4SbCx4Osj1KoVWPcUnH3+rLEyHwiznxgeeUA45VTwgo0rEFewSF5EPC9D7hyCoI1GnMdgQvC41hTuL8q5EKedNWF0iPu+/33Of/f79/+QRoaFSgXnBao9pDYjrsQ4l+XB8iHbHShEQsh1cJlyiAixCRG1RXwGWQJJwvZ77mT1pfvGUozD2B6hWgNivC8IUUiVQ6s2N+jMIUpDNcZKnh3SCYigEbQIRivIMrJNG4jPPOfA7WZWUDr4jikgNmTeI5HB5eG4HsGRUUFTQ6OsQzkBHZEpDRoiLSix4JOQTOD+h+D5L5n1BDGrZYZHv/4l4YQTIY7xtR4kqpC5PM0IgnIpUdbEWEtkM/TYKMY7kFCnQSlIU49YEDE4KyGjw8J5nPGi5/H533735CvBnh1CpQJRBR9VcALVOAbrqChFpD2RZCjbRI2OoJME02wSZ01qeGLxaBEy60hF0RQV0qz0zWP1hc/l0W99db8rUPPRB8QqRdMrbFTBaYPPHForqlpTwWOyJM9z6lFpirKWOA+tMFoRRRoxhiaCNQZfiYhPOJ7hO3+6/5VvcEAwUSAuE5JJ4wwxBp+mRAhKEmKXsEig13t0mqC8C9F+WjCGUCfDOVy9DlmG27yFL33pPzqbCEcJs5oYzr781dBXIzURiTYh6EYrIjxR2iAeGeaxH/0IqjVFtVfRt1AR19Tue++C+giCxcSaLPMhJLMaY7UhiQ09q1Zw9sX75upcd91VQm8vaE0KIUuDzoNelENJCsP9rP/pTyDqUcybr6hVFbWaIq4pv2XTH0Uu/a7OGlQiA0bjjCYRRUibHXHWuRfyn7/zwUknZu2009E9CzCVPjwKiQy+arAuQdk6jc2buOkLn+czv/mbfOIdv8RnPvAB7vvSF5EtWzDWEwlIakNKlTim4S0NDdSqzF+2fNJxbt55i6Ajmk2Lt/mHzuaJ1jJqRhPblN7UUxuok9z+AKM33wlb94SAqdyTVo8lxAhaspASde8evvMf/8kvf/Jzs35XAGavAL3uyq+LNAcky8ZkLA9cqYtIllmRxpgMPPLAAYWzn3z938VJUxpZKt7ltTwSkcQ7aUpTxA3K1uu/u+81moNB8LZWkjygvplJqB+QjsnwY/cemlDYHPmgiH0qCSJo6LuXEDMwNDxpAJPd8bSITSXNg6OKMgxJloo0R+XxG64+YNtbHng0FBlxXrz3MuatDIuTYRERyfYf+zEyIGIzsZIH9ae5kJwmIs0xSTY+Kd/+h4/I+19wibx+6TL5+QVL5R1rjpfXLl0mbzjtVPndt18hn/qzD8m9V18jsnNnUDbs2SK3ff7ThzZWs+SYvTLDjo3CkmVQ6aGJIcssPXFElFrYuQOOnzzUcxzcoHjdi5JKyIGV36qQoBpj7LnlFpZd9vrx18nqEuKZY1LJg/w9KG9JNz5J5bSzD2+Vs1aaURTkSw+9IlAfZWzdY/RdMD4ME5sIKsamHokirIKaAWVTGk8+Ts855x6w7a987BPyi7/9wZA2p1rFKoUNWWDpzZqovXth1dpx19jyox/IcS9/Gak2WCBWGoND2xRGR7nruh/zhU/8Ezs3PY1kjqhWYfVJJ3LGec9mydq1xD09DI812PL0ZtY//gTp6AhvfdMbueisMzn+F94xN3aEHNOQN+kIoVLJg0dGqfX0UlMaMgtpCgP9cPxxBzz9us99TF753vchzqN12PW1AeVdyE02OMzGx55g2WWtc3Y9eKesOPvZiHOgI5RWIfODFhgbZsNdd3HWaWcf3n1o9TXlebt3edJAUdBTo2fpYq76qz+S132oLe17lgCOyFRBHLFS4Z69Jd21i55zDtzU8154SR7spMA5oihCEASFQpMNjRKvGn/OcWecifWgjCEKVUPIyIhdyqP33sOf/8mHmCeac84/n19993s45cILYNmykByhkicV8/ktNBoMbFzPN//zPzj+j/5sThECzGJi2HXDj6ksWIxVMXHffBaee16ZHSIbHiY+wLm3ffVz8sp3vQekB0NIW+kMJBqq4jHOYffs5eLf+W/jHtiKM88Ao8OqjMeJpmoESLB7d3DWO951+A9YufsrmLcDqIIfjwy61kNt4cLx9/zT6zB9C4CI3vkL6Tnt9MCT2yaV5MB5RXfedpOc+twXEtJVeCAK2bwRtBLwQv/OXaw888zxJ65YSUSUB/wrUEF7tHnjZr76ha9wyunP4rd+/3c485IXBeWDJ2QHMQavPVrr0EcF9EUsPvc8fv3P/5Ttt1wvq1/4irlFEDPNp037selxkaRfXDIqNs3EZSKNuhUrImM+Cwa0XZvlrs98cl9+1oW6dIm3kkieG8ylIjImw/cder2JcUc6/NvWZiGQOA+/tlkmMjooO288sAxwKMdtn/9nkR2bRJpjYjMr3rpgGHMujzcOBrn0qafkG3/4P/ZtrzEcEqI5FwyGzop3DXnsrtvkI7/7u7LnkYdC8UGXSGaTIFQUdTSsF5vXfZPEim3mRk8/JmIHZeTBiTX5Zvcxq7VJh4qbv/xxYc+G4Kaw+kRsZSG20oeLI3wEcaww4uj1GQwPcPcPr+KiD/zOJKuWATEYZ4hcqNTjrQIqZHbfXx8SFAvRoTaar4GNwRkFcYRvplO+54ev/Y7gRuT57/plWLECohhjDE4kr7WQUk8GIR3Bb3uaq/7fZ3nL339k33s2GfUoY9gn2MhgndAcS4lqvbzv//sDlp56ImI8aEekXMiBpRMSV0eMxaoM6xoQaXQck4lgVYRXMO+0k9h+x3WzVCidBDNNjZ0cT1z7bZHGDhEZEHF5dgcvMiZORsVJ3TtJkkbQDiUj4revl+//wwFcBNKmZEWp1zzHoUtExHvZe+/hV84REaS594uppCFZhYjUvUgqViQZlf5rvn/Y1xy49SfBzcM1xPlEMpeK9VnpHeG8lTQbFZEREdsvjXX3yb/+1vv23059uzSlIXVxMua8WCuh9K214tJ6cGXJBkWGd4rs3izb77xZNt91s4xteTKUp/WN3A0klSxzknkvTUnFSRL6MLpN7v+vz09t7I7yMeMdmMqx8er/CmVhszGxriENn4bKVpmIT12osCpOxCehjvDgTnn6x4dQ8sgmYrM8d2qRcyivJbXnvnum9kDTAWlKKmPiZUxEGs6JtYnI2JD0X/ntQ75m/103iTT6g+9RlvtB5YTrEydZZqVhE0mykZB2Z+c6uftLh5A/NekX8basOFzmhmkmYpMxkcFd8th3983ROu4Y2haeheT5XTMrmdhAENmg9D80N9ilOcUmXf1//pew62k58WWvCOnXoxqiasQSYyyAoFyGGRxENRswOsLOe++FhSvU8S+7/ODCXNoMqfQVrcqUCnCO+YsWH3Z/dz102yPk1Wp0Xrs00jqI9KOjbFq/b5GUibjv3z8nJEOy+Lzz8dXeUBdB5WlonECSodBEXqi5jMrwMLd9/b9gxanqwl86QNqdAiMJeFOWCG40E3AOjKE+MsIPvvJ1znzjmw98nQWrFZnFZz4ombRBe4OSCHxET8+8QxuwGcacIYanrv+hvPYP/xh6FoEEd4Es8UQWTBNIXcjPkzVgxzao9SnmL1UrX3DpIWs0nnr0ASANah8jLWIQT2XFUjZfc3i1h1ecdMrZSITxmkruLhEBeI8dGGTvnl0HPD99/HE5/xd/FdQ8RM/DeUOSZTiVQU1AN8E3oH8Pw7fdzjUf+RgsWaMuecchOCEWqC4ABQ0FDeUwPZWQw9Vb7NAI25548pAus/Xx9cRKo2y+jnggFVBValQOdvrswExvTYdybLvr9pCdLglaDy8i3rmQZW7Miuwck6d+uJ9sbodxXPP/PimSDIr4wI9b7wLfYDORZEzGnnzwkNvYdMM1OUuTBt4hT9nu00QkGZP6/ipjFseWjSENuwunN23Ijik2EXGjku7eJDd+5V87vuf0gYfFOpExERkWK3XJxCaJiLWS7t4lV3/y44fWxuBQyOCXJ82zBYvZSMU+dOjjNpPH7LVA59h7w42y9AUvAK3wUUTiLSaKsM066x56iPMuvuSwdNn3fetLsmL5UhYpRbpzN4ve/M7x54/uFvpqZComtVCNKkSZDzuFbzK48SkWnXbeAdt8+sfXyQnPez7UQmZyLFDRZOKIjYd0jA0/vpFTXvvzk1/n6UeFlWsR3YOLY9K8ClQ6PMg9N93AS950xUHv+Tt//kfSs2Aevhqz4tSTWL54Ib3Ks+x5r9nnXOdTybTBIBiv0F7hvUVXNDI6zEM/u4lzX72fvgLs6RcWLAwGxViTkGfpdy5Y7h9+kMpFF896m8OsJoYnv/hFOf3n3xwsnVFEIh60YMXiG3U2PvY48ViD5abG0uNWw6pFwQJrc+f6SIU6CjrK6yqo4HyGg+074Lgz9nlA66/+tpz6qkvxpoKlGuIjiPDKY7VgXBMzVmdo/VPc9+Of8PD999PIEtaceBLnXHgR573ghajly0PK/SjOU7oTUmxWNPiU5paNfOPjn+RdH/3UPu3vvedGWXru+UCM11XERKHOgvOM7NxJ/9ObUGMjrFi2mPlrlsP8njDzxICrBDO7s4AnF1KQuIKSjIEnn2DxGftOyrHt66RnxSq0zonXAZHK4yeCQwfpGAz0s+fpp1FKWHryabBwKVANbKuEGhWJz3CVCjEQZylkKbtu/AkrLn9dlxg6wvCgEEXYShUdBWO5eBt89MVhbYaRUOsN8XijQsVNS3AEMjoUWfMh6sYZjTEKJGP08ceZd9azJ39Aw9uFnl4k7kNhwCkyBQ0NCkcfoNMMGhlUY8jrVhMZyCevVwZEE2uCv7/YUJxwbJi7rvo+F79z33pkP/rn/yuXveOXAtHOmwc6j5VAo1U+STMbvjcqCCGSIZFHEYNVrVp2PnfLiDRWPBGO7Q/ey+rz97OTNnYL8XxEV/BOYbTCJRaPI+6pEKo9JoHYvEes4FWMoRL6pQQbZUhkSBFisVS8kD6xjq989OP86v/711lPDLPWHQMI7tNxjESGZpZSMYZIa/AOcZ4orgKUBdktilgkTPi8Wg9F0dEwN0mdoLynOdpgvzqOBasVyaCoRj2UflUV4kjRyC/U9J6aqaBreYVSpSCuhgmBIgMidKhk48NuhnhIE7bcc/ekhABw6rMugp6FYaJ7i8eHqkPO4bwKBQsrIcLNZg4LmLhKsZ7pyCDWhsCivGviQfLIn/ruwf2Pdc9yRTomSjJMFIMoTCXCSETSzNDVCEwfIg6jFKaq81p6+RhXFJFWeDJiHBVvYXc/13/3W3OCEGAWa5Ma990r6Apg8ImlJ65itEbwOGvx3ofZLYZm5mkkFi8KJzov7uFwkj+pvL6HCQslcaRJ66MH7kB1kWJoMFSxqdchzaghRM5R9aCVRorqNXEFLz5vyxIhYBMisSiXhOzejVH23n0Xx73i9fudGKdc/FyIDJl3JOJRxoT6IJHBxFHOAuabUGyoVWNibah4QyyG1FmcAas8GY5MCcpAHBkYrTO8eeuB77nSp5AEfD1YmtM6aKFaiYN2SCDOQ0CxoPJyv0TkWc4TVH2Yik1xW7bw3c/+M6/9X3PHYW/WEkPPCccHfbfzVAVUlqHEoZzDRAZjosAyWEdPFDO/UqFXyJV4AuIw2LyyTihFpZwL6dZ9GlyUD4ZVJ6hs3eNldZ6KtfQ4F+oYW4u4kFrd+QyNBxsIIbYZsVaBd09T5OlNXP2JT7D0xQcJ9yQEycSRoRqF+zOEsE2yJNfNOpTJ2bLUQmLBCkqgx0BNQ+QtsbPELgvE7CwkTXSzcfB7rixQex5+IMgI2kF9JIylViGazubDKQCBTUMnIEkoM5B6Nl5/I+bEZ6k3/dkk7h+zGLOXTdK2VVTQ6yCAOhd47ywLAmqSQN888Fmof+bI43dbRbhxEmrS5jXjlBZwlkp2aH5B8XkXKYAnv/tNOf15z4O+RaF+mxG08eA9Rnzg1V2ozQwGBkfY+9jj3PqjH/H6v/g/6rV/fAhlWv1Y4MkTEzxOTZxHm9mwpbnRXBbIa1RFhpL/sx4lJuwePv9e5bKDE0jrNIeHDumel533QgWw9ZrvyNqXXprXVlB5OKguS1jhmqFQOxl+y05u/eENvOj3/5s66dVXHFI7sw2zW4DuooujiFnLJnXRxdFGlxi66CJHlxi66CJHlxi66CJHlxi66CJHlxi66CJHlxi66CJHlxi66CJHlxi66CJHlxi66CLH/w/Ij8ZjRnVKpgAAAABJRU5ErkJggg==';

  /**
   * The actual tournament emblem image (provided by the organizer),
   * embedded directly so the app has no external image dependency.
   * Shared by the desktop bracket's center column and the mobile round list.
   */
  function trophyHtml() {
    return '<img class="trophy-logo" src="' + LOGO_DATA_URI + '" alt="\u4e16\u5c0fP URA SSP 2026">';
  }

  // ---------------------------------------------------------------------
  // Connector lines (drawn with measured DOM positions, not CSS guesswork)
  // ---------------------------------------------------------------------

  function drawConnectors() {
    var bracketEl = document.getElementById('bracket');
    var svg = document.getElementById('connectors');
    if (!bracketEl || !svg) return;

    // Measure with the mobile zoom transform temporarily neutralized so the
    // path coordinates are in the same natural/untransformed space as the
    // SVG's own CSS box (width/height:100% resolves against the untransformed
    // layout size). The single ancestor transform then scales the whole
    // bracket - cards and connector lines together - so they stay aligned
    // at any zoom level. This happens within one synchronous pass, so there
    // is no visible flash of the un-zoomed layout.
    var prevTransform = bracketEl.style.transform;
    bracketEl.style.transform = 'none';

    // Cards/labels/the trophy carry their own -90deg counter-rotation on
    // mobile (see setCounterRotation_) so they read upright inside the
    // rotated bracket. Neutralize those too while measuring, otherwise their
    // measured boxes would reflect a stray -90deg (parent rotation gone,
    // child's counter-rotation still active) instead of the true layout.
    var cardEls = document.querySelectorAll('.match-card, .round-label, .center-deco');
    var prevCardTransforms = [];
    cardEls.forEach(function (el) {
      prevCardTransforms.push(el.style.transform);
      el.style.transform = 'none';
    });

    var bracketRect = bracketEl.getBoundingClientRect();
    svg.setAttribute('width', bracketRect.width);
    svg.setAttribute('height', bracketRect.height);
    svg.innerHTML = '';

    // Two matches can converge on the same next-match slot from above and
    // below, and their lines cross near that shared card. SVG paints later
    // elements over earlier ones, so drawing strictly in match order lets an
    // undecided line drawn after a confirmed one cut across it. Queue every
    // connector first, then draw all plain lines, then all confirmed (red)
    // ones last, so a red line always ends up on top regardless of which
    // match it belongs to.
    var pending = [];
    matchList.forEach(function (m) {
      if (m.next_match_id) pending.push({ source: m.match_id, target: m.next_match_id, dashed: false, decided: !!m.winner_id });
      if (m.loser_next_match_id) pending.push({ source: m.match_id, target: m.loser_next_match_id, dashed: true, decided: !!m.winner_id });
    });
    pending.sort(function (a, b) { return (a.decided ? 1 : 0) - (b.decided ? 1 : 0); });
    pending.forEach(function (c) { drawConnector(svg, bracketRect, c.source, c.target, c.dashed, c.decided); });

    bracketEl.style.transform = prevTransform;
    cardEls.forEach(function (el, i) { el.style.transform = prevCardTransforms[i]; });
  }

  function drawConnector(svg, bracketRect, sourceId, targetId, dashed, decided) {
    var sourceEl = document.querySelector('.match-card[data-match-id="' + sourceId + '"]');
    var targetEl = document.querySelector('.match-card[data-match-id="' + targetId + '"]');
    if (!sourceEl || !targetEl) return;

    var sRect = sourceEl.getBoundingClientRect();
    var tRect = targetEl.getBoundingClientRect();
    var sourceIsLeftOfTarget = sRect.left < tRect.left;
    var sx = sourceIsLeftOfTarget ? sRect.right : sRect.left;
    var tx = sourceIsLeftOfTarget ? tRect.left : tRect.right;
    var sy = sRect.top + sRect.height / 2;
    var ty = tRect.top + tRect.height / 2;
    var midX = (sx + tx) / 2;

    // The confirmed-winner line pushes past the card edge into the next
    // match's shape, instead of just touching it, so the advancing path
    // reads as "arriving into" that slot. The bend point (midX) stays put -
    // only the final approach into the target lengthens.
    var pierce = (!dashed && decided) ? 16 : 0;
    var pierceTx = sourceIsLeftOfTarget ? (tx + pierce) : (tx - pierce);

    var relSx = sx - bracketRect.left, relTx = pierceTx - bracketRect.left, relMid = midX - bracketRect.left;
    var relSy = sy - bracketRect.top, relTy = ty - bracketRect.top;

    // The winner's advancing line (solid, to next_match_id) turns red once
    // the match result is entered, so the confirmed path through the bracket
    // is easy to trace at a glance. The loser's line to the 3rd-place match
    // stays dashed regardless, since it represents a still-tentative slot.
    var color = dashed ? '#c98f8f' : (decided ? '#FF0002' : '#f1e9e2');

    var path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', 'M ' + relSx + ' ' + relSy + ' H ' + relMid + ' V ' + relTy + ' H ' + relTx);
    path.setAttribute('stroke', color);
    path.setAttribute('stroke-width', (!dashed && decided) ? '6' : '5');
    path.setAttribute('fill', 'none');
    if (dashed) path.setAttribute('stroke-dasharray', '4,3');
    if (!dashed && decided) path.style.filter = 'drop-shadow(0 0 4px rgba(255, 0, 2, 0.75))';
    svg.appendChild(path);
  }

  // ---------------------------------------------------------------------
  // Modal
  // ---------------------------------------------------------------------

  function openModal(matchId) {
    currentMatchId = matchId;
    var m = matchesById[matchId];
    if (!m) return;
    var t1 = teamsById[m.team1_id];
    var t2 = teamsById[m.team2_id];
    var hasTeams = !!(m.team1_id && m.team2_id);

    document.getElementById('modal-title').textContent =
      (t1 ? t1.name : '未定') + ' vs ' + (t2 ? t2.name : '未定');
    document.getElementById('modal-round').textContent =
      (ROUND_LABELS[m.round] || m.round) + ' / ' + matchId;

    document.getElementById('input-date').value = m.date || '';
    document.getElementById('input-time').value = m.time || '';
    document.getElementById('input-venue').value = m.venue || '';
    document.getElementById('input-score1').value = (m.score1 === undefined ? '' : m.score1);
    document.getElementById('input-score2').value = (m.score2 === undefined ? '' : m.score2);

    document.getElementById('score-fields').style.display = hasTeams ? 'flex' : 'none';
    document.getElementById('result-save-btn').style.display = hasTeams ? 'block' : 'none';
    document.getElementById('fusen1-btn').style.display = hasTeams ? 'inline-block' : 'none';
    document.getElementById('fusen2-btn').style.display = hasTeams ? 'inline-block' : 'none';
    document.getElementById('score-disabled-note').hidden = hasTeams;

    document.getElementById('score-team1-label').textContent = t1 ? t1.name : 'チーム1';
    document.getElementById('score-team2-label').textContent = t2 ? t2.name : 'チーム2';
    document.getElementById('fusen1-btn').textContent = (t1 ? t1.name : 'チーム1') + ' 不戦勝';
    document.getElementById('fusen2-btn').textContent = (t2 ? t2.name : 'チーム2') + ' 不戦勝';

    resetClearResultButton_();
    document.getElementById('clear-result-btn').hidden = !m.winner_id;

    document.getElementById('modal-backdrop').hidden = false;
  }

  function closeModal() {
    document.getElementById('modal-backdrop').hidden = true;
    currentMatchId = null;
    resetClearResultButton_();
  }

  function submitSchedule() {
    if (!currentMatchId) return;
    var matchId = currentMatchId;
    var date = document.getElementById('input-date').value;
    var time = document.getElementById('input-time').value;
    var venue = document.getElementById('input-venue').value;

    // Update the bracket from the values just typed instead of waiting on the
    // network round trip, so the modal closes immediately; the real API call
    // still runs in the background and reconciles the view once it resolves.
    var m = matchesById[matchId];
    if (m) Object.assign(m, { date: date, time: time, venue: venue });
    renderAll({ teams: Object.values(teamsById), matches: matchList });
    showToast('日程を保存しました');
    closeModal();

    apiPost_('saveSchedule', { matchId: matchId, date: date, time: time, venue: venue })
      .then(function (data) { renderAll(data); })
      .catch(handleError);
  }

  // Mirrors Code.gs's advanceTeam_: drops teamId into the given match's
  // team1_id/team2_id slot, used for the optimistic client-side update.
  function optimisticAdvance_(nextMatchId, slot, teamId) {
    var m = matchesById[nextMatchId];
    if (!m) return;
    if (String(slot) === '1') m.team1_id = teamId; else m.team2_id = teamId;
  }

  function optimisticApplyResult_(matchId, winnerId, loserId, score1, score2, winType) {
    var m = matchesById[matchId];
    if (!m) return;
    Object.assign(m, { score1: score1, score2: score2, winner_id: winnerId, status: 'completed', win_type: winType });
    if (m.next_match_id) optimisticAdvance_(m.next_match_id, m.next_slot, winnerId);
    if (m.loser_next_match_id) optimisticAdvance_(m.loser_next_match_id, m.loser_next_slot, loserId);
    renderAll({ teams: Object.values(teamsById), matches: matchList });
  }

  function submitResult() {
    if (!currentMatchId) return;
    var matchId = currentMatchId;
    var s1raw = document.getElementById('input-score1').value;
    var s2raw = document.getElementById('input-score2').value;
    if (s1raw === '' || s2raw === '') {
      showToast('得点を入力してください', true);
      return;
    }
    var s1 = Number(s1raw), s2 = Number(s2raw);
    if (isNaN(s1) || isNaN(s2)) {
      showToast('得点は数値で入力してください', true);
      return;
    }
    if (s1 === s2) {
      showToast('同点は保存できません。勝敗が決まってから入力してください。', true);
      return;
    }

    var m = matchesById[matchId];
    var winnerId = (s1 > s2) ? m.team1_id : m.team2_id;
    var loserId = (s1 > s2) ? m.team2_id : m.team1_id;

    // Same instant-close treatment as submitSchedule: apply the winner
    // advancement locally (mirroring the server's logic) and close right
    // away, then reconcile with the real response in the background.
    optimisticApplyResult_(matchId, winnerId, loserId, s1, s2, 'normal');
    showToast('結果を保存しました');
    closeModal();

    apiPost_('saveResult', { matchId: matchId, score1: s1raw, score2: s2raw, winType: 'normal' })
      .then(function (data) { renderAll(data); })
      .catch(handleError);
  }

  function submitFusen(slot) {
    if (!currentMatchId) return;
    var matchId = currentMatchId;
    var winType = (slot === 1) ? 'fusen1' : 'fusen2';
    var m = matchesById[matchId];
    var winnerId = (winType === 'fusen1') ? m.team1_id : m.team2_id;
    var loserId = (winType === 'fusen1') ? m.team2_id : m.team1_id;

    optimisticApplyResult_(matchId, winnerId, loserId, '', '', winType);
    showToast('不戦勝を記録しました');
    closeModal();

    apiPost_('saveResult', { matchId: matchId, score1: '', score2: '', winType: winType })
      .then(function (data) { renderAll(data); })
      .catch(handleError);
  }

  var clearConfirmTimer = null;

  function resetClearResultButton_() {
    clearTimeout(clearConfirmTimer);
    var btn = document.getElementById('clear-result-btn');
    btn.dataset.confirming = '';
    btn.textContent = '結果の取り消し';
  }

  // Mirrors Code.gs's retractTeam_: clears the given match's team1_id/team2_id
  // slot, used for the optimistic client-side update.
  function optimisticRetract_(nextMatchId, slot) {
    var m = matchesById[nextMatchId];
    if (!m) return;
    if (String(slot) === '1') m.team1_id = ''; else m.team2_id = '';
  }

  function submitClearResult() {
    if (!currentMatchId) return;
    var matchId = currentMatchId;
    var btn = document.getElementById('clear-result-btn');

    if (!btn.dataset.confirming) {
      btn.dataset.confirming = '1';
      btn.textContent = '本当に取り消しますか？（もう一度クリック）';
      clearTimeout(clearConfirmTimer);
      clearConfirmTimer = setTimeout(function () { resetClearResultButton_(); }, 4000);
      return;
    }
    resetClearResultButton_();

    // Same check as Code.gs's assertNextMatchNotDecided_: refuse to clear if
    // the match(es) this one feeds into are already decided, since retracting
    // the team here would leave that later result pointing at a team that
    // never actually qualified.
    var m = matchesById[matchId];
    var blockedBy = [m.next_match_id, m.loser_next_match_id].filter(function (id) {
      return id && matchesById[id] && matchesById[id].winner_id;
    })[0];
    if (blockedBy) {
      showToast('進出先の試合（' + blockedBy + '）の結果が既に入力されているため取り消せません。先にそちらの結果を取り消してください。', true);
      return;
    }

    Object.assign(m, { score1: '', score2: '', winner_id: '', status: '', win_type: '' });
    if (m.next_match_id) optimisticRetract_(m.next_match_id, m.next_slot);
    if (m.loser_next_match_id) optimisticRetract_(m.loser_next_match_id, m.loser_next_slot);
    renderAll({ teams: Object.values(teamsById), matches: matchList });
    showToast('結果を取り消しました');
    closeModal();

    apiPost_('clearResult', { matchId: matchId })
      .then(function (data) { renderAll(data); })
      .catch(handleError);
  }

  // ---------------------------------------------------------------------
  // Team menu (hamburger) + team icon upload
  // ---------------------------------------------------------------------

  function openTeamMenu() {
    renderTeamMenuList_();
    document.getElementById('team-menu-backdrop').hidden = false;
  }

  function closeTeamMenu() {
    document.getElementById('team-menu-backdrop').hidden = true;
  }

  /**
   * Teams sorted in Japanese (a-i-u-e-o) reading order. Name is
   * display-only here - it can't be edited, only the icon can.
   */
  function renderTeamMenuList_() {
    var container = document.getElementById('team-menu-list');
    if (!container) return;

    var teams = Object.keys(teamsById).map(function (id) { return teamsById[id]; });
    teams.sort(function (a, b) { return (a.kana || a.name || '').localeCompare(b.kana || b.name || '', 'ja'); });

    container.innerHTML = teams.map(function (t) {
      return '<button type="button" class="team-menu-row" onclick="openIconModalForTeam(\'' + t.team_id + '\')">' +
        badgeHtml(t) +
        '<span class="team-menu-name">' + escapeHtml(t.name) + '</span>' +
        '</button>';
    }).join('');
  }

  function openIconModalForTeam(teamId) {
    var team = teamsById[teamId];
    if (!team) return;
    currentIconTeamId = teamId;
    document.getElementById('icon-modal-team-name').textContent = team.name;
    document.getElementById('icon-modal-preview').innerHTML = badgeHtml(team);
    document.getElementById('icon-modal-backdrop').hidden = false;
  }

  function closeIconModal() {
    document.getElementById('icon-modal-backdrop').hidden = true;
    currentIconTeamId = null;
  }

  function triggerIconUploadForCurrentTeam() {
    if (!currentIconTeamId) return;
    document.getElementById('icon-modal-file').click();
  }

  function onIconModalFileSelected(event) {
    var file = event.target.files && event.target.files[0];
    event.target.value = ''; // allow re-selecting the same file later
    if (!file) return;
    var teamId = currentIconTeamId;
    if (!teamId) return;

    if (file.size > 5 * 1024 * 1024) {
      showToast('画像サイズは5MB以内にしてください', true);
      return;
    }

    resizeImageForBadge_(file)
      .then(function (dataUrl) {
        // Same instant-feedback treatment as the other save actions: apply
        // the resized image locally and update the UI immediately instead
        // of waiting on the network round trip (which also uploads a Drive
        // backup copy server-side, adding real latency) - the actual save
        // still happens right after, in the background.
        var team = teamsById[teamId];
        if (team) team.logo_url = dataUrl;
        renderAll({ teams: Object.values(teamsById), matches: matchList });
        refreshIconModalAndMenu_(teamId);
        showToast('アイコンを変更しました');

        return apiPost_('uploadTeamLogo', { teamId: teamId, dataUrl: dataUrl, fileName: file.name });
      })
      .then(function (data) {
        renderAll(data);
        refreshIconModalAndMenu_(teamId);
      })
      .catch(handleError);
  }

  /**
   * Shrinks the picked image to a small square-ish JPEG data URI before
   * uploading. This isn't about format compatibility - it's because
   * uploadTeamLogo now stores the image data directly in the Teams sheet
   * cell (see the comment in Code.gs for why: Google Drive's own share
   * links don't reliably load in a plain <img> tag), and a sheet cell only
   * holds ~50,000 characters. A badge is only ever shown at a few dozen
   * pixels across, so shrinking well below the original resolution costs
   * nothing visually.
   */
  function resizeImageForBadge_(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onload = function () {
        var img = new Image();
        img.onload = function () {
          var maxSize = 240;
          var scale = Math.min(1, maxSize / Math.max(img.naturalWidth, img.naturalHeight));
          var canvas = document.createElement('canvas');
          canvas.width = Math.round(img.naturalWidth * scale) || img.naturalWidth;
          canvas.height = Math.round(img.naturalHeight * scale) || img.naturalHeight;
          canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
          try {
            resolve(canvas.toDataURL('image/jpeg', 0.82));
          } catch (e) {
            reject(e);
          }
        };
        img.onerror = function () {
          reject(new Error('画像を読み込めませんでした。別の画像でお試しください。'));
        };
        img.src = reader.result;
      };
      reader.onerror = function () {
        reject(new Error('画像の読み込みに失敗しました'));
      };
      reader.readAsDataURL(file);
    });
  }

  function resetCurrentTeamIcon() {
    var teamId = currentIconTeamId;
    if (!teamId) return;

    var team = teamsById[teamId];
    if (team) team.logo_url = '';
    renderAll({ teams: Object.values(teamsById), matches: matchList });
    refreshIconModalAndMenu_(teamId);
    showToast('アイコンを元に戻しました');

    apiPost_('clearTeamLogo', { teamId: teamId })
      .then(function (data) {
        renderAll(data);
        refreshIconModalAndMenu_(teamId);
      })
      .catch(handleError);
  }

  function refreshIconModalAndMenu_(teamId) {
    var team = teamsById[teamId];
    if (!team) return;
    var preview = document.getElementById('icon-modal-preview');
    if (preview) preview.innerHTML = badgeHtml(team);
    if (!document.getElementById('team-menu-backdrop').hidden) renderTeamMenuList_();
  }

  // ---------------------------------------------------------------------
  // Utilities
  // ---------------------------------------------------------------------

  function handleError(err) {
    showToast((err && err.message) ? err.message : String(err), true);
  }

  function showToast(msg, isError) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.hidden = false;
    t.className = 'toast' + (isError ? ' error' : '');
    clearTimeout(showToast._timer);
    showToast._timer = setTimeout(function () { t.hidden = true; }, 3000);
  }

  var WEEKDAY_LABELS = ['日', '月', '火', '水', '木', '金', '土'];

  function formatDateLabel(dateStr) {
    var parts = String(dateStr).split('-');
    if (parts.length !== 3) return dateStr;
    var y = Number(parts[0]), mo = Number(parts[1]), d = Number(parts[2]);
    var dt = new Date(y, mo - 1, d);
    var weekday = WEEKDAY_LABELS[dt.getDay()];
    return mo + '/' + d + '(' + weekday + ')';
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function debounce(fn, wait) {
    var timer;
    return function () {
      clearTimeout(timer);
      timer = setTimeout(fn, wait);
    };
  }
