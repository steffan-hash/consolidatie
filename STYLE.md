# Stijlgids — Pallet Consolidatie

Overdraagbare stijlreferentie, bedoeld om mee te nemen naar een ander
project. Beschrijft de visuele stijl van deze tool: waarom de keuzes zijn
gemaakt (zie ook `Log.md`, sessie 24-08), en de concrete Tailwind-classes om
over te nemen.

## Fundament

- **Tailwind CSS via CDN** (`<script src="https://cdn.tailwindcss.com">`) —
  geen build-stap nodig, geschikt voor een project zonder toolchain (losse
  `.html`/`.js`-bestanden, direct te openen of te hosten).
- **Lettertype:** Inter (Google Fonts), fallback `Segoe UI, Arial, sans-serif`.
- **Stijlrichting:** geïnspireerd op [ui.shadcn.com](https://ui.shadcn.com) —
  sober, neutrale grijstinten, kleine/strakke afgeronde hoeken, compacte
  knoppen. Geen zware schaduwen of felle kleuren.
- **Kleurgebruik:** Tailwind's ingebouwde `zinc`-grijsschaal voor alles
  behalve accenten — geen eigen grijstinten gedefinieerd. Eén accentkleur
  (hier geel), alleen gebruikt op knoppen, checkboxen en focus-states — niet
  overal verspreid.
- **Donker thema:** `darkMode: 'class'` (niet automatisch op
  systeeminstelling varen, maar expliciet via een class op `<html>`, zodat
  een knop het kan omschakelen). Voorkeur onthouden in `localStorage`; bij
  een eerste bezoek de systeeminstelling volgen. Wordt vóór de rest van de
  pagina toegepast (blocking script in `<head>`, vóór Tailwind laadt) om een
  lichtflits te voorkomen bij het openen in donker thema.

## HTML-skelet (in `<head>`)

```html
<!-- Thema zo vroeg mogelijk toepassen, vóór de rest van de pagina rendert -->
<script>
  (function () {
    var stored = localStorage.getItem('theme');
    var prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    var dark = stored ? stored === 'dark' : prefersDark;
    if (dark) document.documentElement.classList.add('dark');
    document.documentElement.style.colorScheme = dark ? 'dark' : 'light';
  })();
</script>

<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">

<script src="https://cdn.tailwindcss.com"></script>
<script>
  tailwind.config = {
    darkMode: 'class',
    theme: {
      extend: {
        fontFamily: { sans: ['Inter', 'Segoe UI', 'Arial', 'sans-serif'] },
        colors: {
          accent: '#f7c948',
          accentdark: '#eab627',
        },
      },
    },
  }
</script>
```

Pas de twee accentkleuren aan naar de merkkleur van het nieuwe project — de
rest van de stijl (grijstinten, hoeken, spacing) staat er los van.

## Thema-knop (licht/donker wisselen)

```html
<button
  type="button"
  id="themeToggle"
  aria-label="Wissel tussen licht en donker thema"
  class="inline-flex items-center justify-center w-9 h-9 shrink-0 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 hover:bg-zinc-50 dark:hover:bg-zinc-700 transition-colors cursor-pointer"
>
  <svg id="iconSun" class="hidden w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></svg>
  <svg id="iconMoon" class="w-4 h-4" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79Z"/></svg>
</button>
```

```js
themeToggle.addEventListener('click', () => {
  const isDark = document.documentElement.classList.toggle('dark');
  document.documentElement.style.colorScheme = isDark ? 'dark' : 'light';
  localStorage.setItem('theme', isDark ? 'dark' : 'light');
});
```

## Terugkerende classes

**Pagina-achtergrond en basistekst**
```
bg-zinc-50 dark:bg-zinc-950 text-zinc-900 dark:text-zinc-100 font-sans leading-relaxed
```

**Pagina-breedte**
```
max-w-screen-2xl mx-auto px-6 py-10
```
(Gebruik een smallere waarde als het geen brede datatabel betreft, bv. `max-w-4xl`.)

**Kaart** (de basisbouwsteen van de hele UI)
```
bg-white dark:bg-zinc-900 border border-zinc-200 dark:border-zinc-800 rounded-xl p-6 shadow-sm
```

**Titel (H1) / kaartkop (H2)**
```
H1: text-2xl font-bold tracking-tight
H2: text-base font-semibold mb-4
Subtekst onder H1: text-zinc-500 dark:text-zinc-400 text-sm mt-1
```

**Primaire knop** (accentkleur, belangrijkste actie)
```
inline-flex items-center justify-center h-9 bg-accent hover:bg-accentdark
text-zinc-900 font-medium text-sm px-4 rounded-md transition-colors cursor-pointer
```
Disabled-variant erbij: `disabled:bg-zinc-100 dark:disabled:bg-zinc-800 disabled:text-zinc-400 dark:disabled:text-zinc-500 disabled:cursor-not-allowed`

**Secundaire/icoon-knop** (bv. thema-toggle, sluitknop)
```
inline-flex items-center justify-center rounded-md border border-zinc-200 dark:border-zinc-700
bg-white dark:bg-zinc-800 text-zinc-600 dark:text-zinc-300 hover:bg-zinc-50 dark:hover:bg-zinc-700
transition-colors cursor-pointer
```

**Tekstinput**
```
w-full h-9 rounded-md border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-800
text-sm px-3 focus:outline-none focus:ring-2 focus:ring-accent
```

**Checkbox**
```
w-4 h-4 rounded border-zinc-300 dark:border-zinc-600 accent-[#eab627]
```
(De inline hex hier mag ook `accent-accentdark` worden als die kleur in de
Tailwind-config staat — hier bewust los gehouden van de config-naam.)

**Kleine statistiektegel** (getal + label, in een grid)
```html
<div class="bg-white dark:bg-zinc-800 border border-zinc-200 dark:border-zinc-700 rounded-lg px-3.5 py-3">
  <div class="text-xl font-semibold">{getal}</div>
  <div class="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">{label}</div>
</div>
```
Grid eromheen: `grid grid-cols-[repeat(auto-fit,minmax(140px,1fr))] gap-3`

**Tabel** (koprij vastgezet, hover per rij)
```
Kop (th):  sticky top-0 z-10 bg-zinc-50 dark:bg-zinc-800 text-left font-medium
           text-xs uppercase tracking-wide text-zinc-500 dark:text-zinc-400
           px-3 py-2 border-b border-zinc-200 dark:border-zinc-700 whitespace-nowrap
Cel (td):  px-3 py-2 border-b border-zinc-100 dark:border-zinc-800 whitespace-nowrap
Rij (tr):  hover:bg-zinc-50 dark:hover:bg-zinc-800/60
Wrapper:   max-h-[420px] overflow-auto border border-zinc-200 dark:border-zinc-700 rounded-lg
```

**Melding/hint-balk** (waarschuwing of fout, boven de inhoud)
```
Fout:        bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-900 text-red-700 dark:text-red-300
Waarschuwing: bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-900 text-amber-800 dark:text-amber-300
Beide:       px-4 py-3 rounded-lg text-sm
```

**Uitgelicht/opvallend tekst in een tabel** (bv. een status die aandacht vraagt)
```
font-medium text-[#8a6d1a] dark:text-[#eab627]
```
(Zelfde geel-familie als de accentkleur, maar met eigen licht/donker-tint
voor voldoende contrast — niet dezelfde hex in beide thema's.)

## Vuistregels

- Kleine, strakke hoeken: `rounded-md`/`rounded-lg`/`rounded-xl`. Geen
  `rounded-2xl` of `rounded-full` (behalve bij een puur ronde iconknop).
- Knoppen vast op `h-9`, tekst `text-sm font-medium` — geen dikke padding of
  vetgedrukte (`font-bold`) knoptekst.
- Tabelkoppen klein en ingetogen: `text-xs uppercase tracking-wide` met
  gedempte kleur (`text-zinc-500`), niet de volle tekstkleur.
- Altijd een `dark:`-variant meegeven bij elke kleur — nooit alleen een
  lichte kleur zetten en het donkere thema laten "meeliften" op een
  toevallig passende tint.
- Layout met flexbox/grid en `gap`, niet met losse margins tussen
  elementen (voorkomt dubbele/wegvallende marges).
