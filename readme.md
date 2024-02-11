
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>
<br/>

<h3 align='center'>@lancejpollard/md2docx.js</h3>
<p align='center'>
  Mardown to DOCX in TypeScript
</p>

<br/>
<br/>
<br/>

## Status

This is a new project and I [need some help](https://github.com/dolanmiu/docx/discussions/2589) with styling it to look nice, haven't fully figured out the `docx` API. But it is a pretty good start!

## Installation

```
pnpm install @lancejpollard/md2docx
```

## Usage

```ts
import md2docx from '@lancejpollard/md2docx'
import fs from 'fs'

md2docx(
  `# Hello World

I am some text.

- a list item
- a second list item

I am a [link](https://google.com). And _italics_ and **bold**.

_**bold with italics**_.

- nested list
  - list item
  - numbered
    1. one
    2. two

1. a
2. b

- **some bold _italics_**

## **h2 _italics_** and more

`,
  'buffer',
).then(buffer => {
  fs.writeFileSync('test/md.docx', buffer)
})
```

## License

MIT
