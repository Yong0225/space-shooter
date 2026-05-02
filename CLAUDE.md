# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the game

Open `shooter.html` directly in a browser — no build step, no server required. Double-click the file or drag it into a browser tab.

## Architecture

The entire game lives in a single file: `shooter.html`. It is structured in three sections inline:

- **CSS** (`<style>`) — layout, overlay, HUD, and button styles. Canvas is fixed at 480×600.
- **HTML** — a HUD bar (`#ui`), the `<canvas id="c">`, and an `#overlay` div reused for start/game-over screens.
- **JavaScript** (`<script>`) — all game logic, no external dependencies.

### Game loop

`loop(ts)` is driven by `requestAnimationFrame`. Each tick computes `dt` normalised to 60 fps (`ts / 16.67`, clamped to 3), then calls `update(dt)` → `render()`.

### State machine

The global `state` string controls what `update` and `render` act on:

| Value | Meaning |
|-------|---------|
| `idle` | Title screen, loop not yet started |
| `playing` | Active gameplay |
| `dead` | Game over overlay shown |

### Key subsystems

| Function | Responsibility |
|----------|---------------|
| `initGame()` | Resets all globals and calls `buildEnemyGrid(1)` |
| `buildEnemyGrid(lvl)` | Rebuilds the `enemies` array; rows = `min(2+lvl, 5)`, cols = 8 |
| `updateEnemies(dt)` | Marches the grid sideways, drops a row on wall-hit, fires enemy bullets, scales speed by survivors remaining |
| `triggerDeath()` | Decrements lives, grants 180-frame invincibility, or transitions to `dead` |
| `rectsOverlap()` | AABB collision used for all bullet/entity hits |
| `render()` | Full repaint every frame: background → stars → particles → player → enemies → bullets |
| `showOverlay()` | Rebuilds `#overlay` innerHTML and re-attaches the start button listener |

### Difficulty scaling (per level)

- Enemy grid grows by one row per level (max 5 rows).
- Enemy move speed increases as fewer enemies remain: `speed = max(4, 40 - aliveCount * 0.5)`.
- Enemy shoot interval shrinks: `max(30, 90 - level * 8)` frames.
- Enemy bullet speed increases: `4 + level * 0.4` px/frame.
- Score per kill multiplied by level.

### Data shapes

```js
player      = { x, y, w, h, speed, invincible }   // invincible counts down in dt units
enemy       = { x, y, w, h, color, hp, alive, col, row }
bullet      = { x, y, speed }                       // player bullets move up
enemyBullet = { x, y, speed }                       // enemy bullets move down
particle    = { x, y, vx, vy, life, decay, size, color }
```

## Git workflow

Remote: https://github.com/Yong0225/space-shooter  
Branch: `master`

Commit message convention: `type: short description` (e.g. `feat:`, `fix:`, `refactor:`).  
After each meaningful change: `git add shooter.html && git commit -m "..." && git push`.
