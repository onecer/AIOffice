# animations (pptx)

Add an animation to a shape, then retime or reorder it. The op path targets the
shape: `/slide[i]/shape[@id=N]`.

```jsonc
{op:"add", path:"/slide[1]/shape[@id=5]", type:"animation",
 props:{effect:"fade", trigger:"click", duration:500, delay:0}}
```

## effects

| class | effects |
|-------|---------|
| entrance | `appear` · `fade` · `flyIn` · `wipe` |
| emphasis | `pulse` · `grow` · `spin` · `colorPulse` |
| exit | `fadeOut` · `flyOut` · `wipeOut` |
| motion (**1.3**) | `motionPath` |

`flyIn`/`flyOut`/`wipe`/`wipeOut` take a `direction` (`left`/`right`/`top`/
`bottom`); `colorPulse` takes a `color`. `trigger`: `click` (default) ·
`afterPrevious` · `withPrevious`.

## click triggers (1.4)

`triggerOn:"@N"` makes the effect play when **another shape** (stable id `N`) is
clicked, instead of advancing in the main click sequence. The effect joins that
trigger shape's interactive `onClick` sequence.

```jsonc
{op:"add", path:"/slide[1]/shape[@id=5]", type:"animation",
 props:{effect:"fade", triggerOn:"@9"}}   // shape 5 fades in when shape 9 is clicked
```

- `triggerOn` names a **different** shape on the same slide (a shape cannot trigger
  its own animation); an unknown id is `invalid_path` with the available ids.
- `read --view structure` and `get /slide[i]/animation[k]` report the trigger shape.
- Works with every effect class, including `motionPath`. `validate` stays clean.

## repeat / rewind / auto-reverse (1.7)

Add the animation, then set its timing on the animation path
`set /slide[i]/animation[k]` (timing is a retime, not an add prop):

```jsonc
{op:"add", path:"/slide[1]/shape[@id=5]", type:"animation", props:{effect:"pulse"}}
{op:"set", path:"/slide[1]/animation[1]", props:{repeat:"untilClick", autoReverse:true, rewind:false}}
```

- `repeat`: `"none"` (play once, default) · a count `1..10000` (loop N times) ·
  `"untilClick"` / `"untilNext"` (loop until the next click). Stored as the
  effect's `repeatCount` (OOXML thousandths-of-an-iteration; `"indefinite"` for
  the loop forms).
- `rewind` (bool): PowerPoint's "Rewind when done playing" — `true` returns the
  shape to its pre-animation state (`fill="remove"`); `false` (default) holds the
  final state (`fill="hold"`).
- `autoReverse` (bool): play the effect, then play it backwards.

`read --view structure` and `get /slide[i]/animation[k]` report `repeat`,
`rewind` and `autoReverse`. All combinations validate OpenXmlValidator-clean.

## motion paths (1.3)

`effect:"motionPath"` moves the shape along a path instead of animating in place.

```jsonc
{op:"add", path:"/slide[1]/shape[@id=5]", type:"animation",
 props:{effect:"motionPath", path:"arc", trigger:"withPrevious", duration:1000}}
```

- `path`: `line` · `arc` · `circle` · `custom`.
- `direction` orients `line`/`arc`.
- `custom` takes a `points` array of normalised `[x,y]` pairs (0..1 of the slide)
  to trace your own path.

`read --view structure` lists every animation (with its effect, including
`motionPath`) on a slide. `set /slide[i]/animation[k]` retimes `trigger`/`delay`/
`duration` in place. All combinations validate OpenXmlValidator-clean.
