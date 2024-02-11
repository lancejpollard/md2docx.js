import md2docx from '..'
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

Then this:

1. a
2. b

- **some bold _italics_**

## **h2 _italics_** and more

`,
  { output: 'buffer' },
).then(buffer => {
  if (buffer instanceof Buffer) {
    fs.writeFileSync('test/md.docx', buffer)
  }
})
