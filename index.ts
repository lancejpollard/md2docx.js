import { marked } from 'marked'
import {
  Packer,
  Paragraph,
  HeadingLevel,
  Document,
  TextRun,
  ExternalHyperlink,
  AlignmentType,
  Run,
} from 'docx'

export type Encoding = 'buffer' | 'blob' | 'base64'

export default async function md2docx<E extends Encoding>(
  text: string,
  encoding: E,
) {
  const document: Array<any> = []

  const HEADING: Record<number, any> = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6,
  }

  // Override function
  const walkTokens = token => {
    switch (token.type) {
      case 'space':
        document.push(new TextRun({ text: token.raw }))
        break
      case 'list':
        document.push(
          ...walkList(token, { depth: 0, ordered: token.ordered }),
        )
        break
      case 'heading':
        document.push(walkHeading(token))
        break
      case 'paragraph': {
        const children: Array<any> = []
        token.tokens.forEach(child => {
          children.push(
            ...walkText(child, { italics: false, bold: false }),
          )
        })
        children.push(new Run({ break: 1 }))
        const p = new Paragraph({ children })
        document.push(p)
        break
      }
    }
  }

  const walkHeading = token => {
    const heading = HEADING[token.depth]
    const children: Array<any> = []

    token.tokens.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push(new TextRun({ text: child.text }))
          break
        case 'link':
          children.push(
            walkLink(token, { italics: false, bold: false }),
          )
          break
        case 'em':
          children.push(...walkEm(token, { bold: false }))
          break
        case 'strong':
          children.push(...walkStrong(token, { italics: false }))
          break
      }
    })

    return new Paragraph({ heading, children })
  }

  const walkStrong = (token, { italics }) => {
    const children: Array<any> = []

    token.tokens.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push(
            new TextRun({ text: child.text, bold: true, italics }),
          )
          break
        case 'link':
          children.push(walkLink(token, { italics, bold: true }))
          break
        case 'em':
          children.push(...walkEm(token, { bold: true }))
          break
      }
    })

    return children
  }

  const walkEm = (token, { bold }) => {
    const children: Array<any> = []

    token.tokens.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push(
            new TextRun({ text: child.text, bold, italics: true }),
          )
          break
        case 'link':
          children.push(walkLink(token, { italics: true, bold }))
          break
        case 'strong':
          children.push(...walkStrong(token, { italics: true }))
          break
      }
    })

    return children
  }

  const walkLink = (token, { italics, bold }) => {
    const children: Array<any> = []
    token.tokens.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push(
            new TextRun({ text: child.text, italics, bold }),
          )
          break
        case 'em':
          children.push(...walkEm(token, { bold }))
          break
        case 'strong':
          children.push(...walkStrong(token, { italics }))
          break
      }
    })
    return new ExternalHyperlink({ link: token.href, children })
  }

  const walkText = (token, { italics, bold }) => {
    const children: Array<any> = []
    if (token.tokens?.length) {
      token.tokens.forEach(child => {
        switch (child.type) {
          case 'text':
            children.push(new TextRun({ text: child.text }))
            break
          case 'link':
            children.push(walkLink(token, { italics, bold }))
            break
          case 'em':
            children.push(...walkEm(token, { bold }))
            break
          case 'strong':
            children.push(...walkStrong(token, { italics }))
            break
        }
      })
    } else {
      children.push(new TextRun({ text: token.text }))
    }
    return children
  }

  const walkList = (token, { ordered, depth = 0 }) => {
    const children: Array<any> = []
    token.items.forEach(item => {
      children.push(walkListItem(item, { ordered, depth }))
    })
    if (depth === 0) {
      children.push(new Run({ break: 2 }))
    }
    return children
  }

  const walkListItem = (token, { ordered, depth = 0 }) => {
    const opts = ordered
      ? {
          numbering: {
            reference: `my-crazy-numbering`,
            level: depth,
          },
        }
      : {
          bullet: {
            level: depth,
          },
        }
    const children: Array<any> = []

    token.tokens.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push(
            ...walkText(child, { italics: false, bold: false }),
          )
          break
        case 'list':
          children.push(
            ...walkList(child, {
              ordered: child.ordered,
              depth: depth + 1,
            }),
          )
          break
      }
    })

    return new Paragraph({ ...opts, children })
  }

  marked.use({ walkTokens })

  marked.parse(text)

  const doc = new Document({
    title: 'Sample Document',
    description: 'A brief example of using docx',
    numbering: {
      config: [
        {
          reference: 'my-crazy-numbering',
          levels: [
            {
              level: 0,
              format: 'decimal',
              text: '%1.',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: { left: 720, hanging: 260 },
                },
              },
            },
            {
              level: 1,
              format: 'lowerLetter',
              text: '%2)',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: { left: 1440, hanging: 980 },
                },
              },
            },
            {
              level: 2,
              format: 'upperLetter',
              text: '%3)',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: { left: 2160, hanging: 1700 },
                },
              },
            },
            {
              level: 3,
              format: 'decimal',
              text: '%4.',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: { left: 2880, hanging: 2420 },
                },
              },
            },
          ],
        },
      ],
    },
    styles: {
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: 28,
            bold: true,
            italics: true,
            color: '#ff0000',
          },
          paragraph: {
            spacing: {
              after: 120,
            },
          },
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: 26,
            bold: true,
          },
          paragraph: {
            spacing: {
              before: 240,
              after: 120,
            },
          },
        },
      ],
    },
    sections: [
      {
        children: document,
      },
    ],
  })

  switch (encoding) {
    case 'buffer':
      return await Packer.toBuffer(doc)
    case 'blob':
      return await Packer.toBlob(doc)
  }

  return await Packer.toBase64String(doc)
}
