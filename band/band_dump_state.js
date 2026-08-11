/* band_dump_state.js - ask the live page for the collector's own state.
 *
 * Downloads __grabstate__.json so a PowerShell/Python caller can read it back.
 * band/liveness.py polls this; inject_here.ps1 pastes it (proving the origin first).
 *
 * WHY IT ALSO READS localStorage (incident #35, 2026-08-11):
 *   When the TOP document navigates, the injected collector dies instantly and
 *   window.__GRAB is gone. The old probe could only say "NO __GRAB", which reads
 *   the same as "was never injected" - and several sessions did read it that way,
 *   concluding "comments not scraped yet" when the truth was "scraping died".
 *   grab_posts.js now writes a heartbeat (__grabBeat) and a death record
 *   (__grabDeath) into localStorage. Those SURVIVE the navigation because the
 *   origin is still band.us, so this probe can tell the three apart:
 *       never injected  -> no global, no beat, no death
 *       died            -> no global, but a beat/death record with counts
 *       alive           -> global answers
 *
 * The literal string NO __GRAB is a contract: browser_chain.py and
 * inject_and_verify.ps1 both grep for it as the death signature. Do not rename it.
 * ASCII-only on purpose - this text travels through the Windows clipboard.
 */
(function () {
  var out = {};
  function ls(k) { try { return JSON.parse(localStorage.getItem(k) || 'null'); } catch (e) { return null; } }
  var beat = ls('__grabBeat'), death = ls('__grabDeath');
  try {
    var s = (typeof window.__grabStatus === 'function') ? window.__grabStatus() : window.__GRAB;
    if (s === undefined || s === null) {
      // Keep the exact legacy marker, then say what actually happened.
      out.err = 'NO __GRAB - script is dead (page navigated?)';
      out.everRan = !!(beat || death);
      out.verdict = (beat || death) ? 'DIED_AFTER_START' : 'NEVER_STARTED';
    } else {
      out.status = s;
      out.verdict = 'ALIVE';
    }
  } catch (e) {
    out.err = String(e && e.message);
    out.verdict = 'PROBE_ERROR';
  }
  out.beat = beat;
  out.death = death;
  out.href = location.href;
  out.title = document.title;
  out.hidden = !!document.hidden;
  var j;
  try { j = JSON.stringify(out, null, 1); } catch (e) { j = '{"err":"stringify"}'; }
  if (j.length > 4000) { j = j.slice(0, 4000); }
  var a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([j], { type: 'application/json' }));
  a.download = '__grabstate__.json';
  document.body.appendChild(a);
  a.click();
  a.remove();
  return 'grab state ' + j.length;
})()
