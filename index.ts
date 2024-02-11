import {
  Packer,
  Paragraph,
  HeadingLevel,
  Document,
  TextRun,
  ExternalHyperlink,
  AlignmentType,
  Run,
  convertInchesToTwip,
  LevelFormat,
} from 'docx'
import { fromMarkdown } from 'mdast-util-from-markdown'
// https://github.com/inokawa/remark-docx/blob/main/src/plugin.ts

// new ImageRun({
//   data: Buffer.from(imageBase64Data, 'base64'),
//   transformation: {
//     width: 100,
//     height: 100,
//   },
// })

export type Output = 'buffer' | 'blob' | 'base64'

export type Options = {
  output: Output
}

const INDENT = 0.17

const HEADING: Record<number, any> = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
}

export default async function md2docx(
  text: string,
  options: Options = { output: 'buffer' },
) {
  const node = fromMarkdown(text)
  const json = walk(node)
  const children = json.map(child => convert(child))

  // console.log(JSON.stringify(json, null, 2))
  // return

  const doc = new Document({
    title: 'Sample Document',
    description: 'A brief example of using docx',
    numbering: {
      config: [
        {
          reference: 'bullet',
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  // spacing: {
                  //   before: 100,
                  // },
                  indent: {
                    left: convertInchesToTwip(INDENT * 0),
                  },
                },
              },
            },
            {
              level: 1,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 1),
                  },
                },
              },
            },
            {
              level: 2,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 2),
                  },
                },
              },
            },
            {
              level: 3,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 3),
                  },
                },
              },
            },
            {
              level: 4,
              format: LevelFormat.BULLET,
              text: '•',
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 4),
                  },
                },
              },
            },
          ],
        },
        {
          reference: 'number',
          levels: [
            {
              level: 0,
              format: 'decimal',
              text: '%1.',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 0),
                  },
                },
              },
            },
            {
              level: 1,
              format: 'decimal',
              text: '%2.',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 1),
                  },
                },
              },
            },
            {
              level: 2,
              format: 'decimal',
              text: '%3.',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 2),
                  },
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
                  indent: {
                    left: convertInchesToTwip(INDENT * 3),
                  },
                },
              },
            },
            {
              level: 4,
              format: 'decimal',
              text: '%5.',
              alignment: AlignmentType.START,
              style: {
                paragraph: {
                  indent: {
                    left: convertInchesToTwip(INDENT * 4),
                  },
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
            size: '32pt',
            // bold: true,
            // italics: true,
            // color: '#ff0000',
          },
          paragraph: {
            spacing: {
              after: convertInchesToTwip(INDENT * 1),
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
            size: '28pt',
            // bold: true,
          },
          paragraph: {
            spacing: {
              before: convertInchesToTwip(INDENT * 2),
              after: convertInchesToTwip(INDENT * 1),
            },
          },
        },
        {
          id: 'Heading3',
          name: 'Heading 3',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: '24pt',
          },
          paragraph: {
            spacing: {
              before: convertInchesToTwip(INDENT * 2),
              after: convertInchesToTwip(INDENT * 1),
            },
          },
        },
        {
          id: 'Heading4',
          name: 'Heading 4',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: '20pt',
          },
          paragraph: {
            spacing: {
              before: convertInchesToTwip(INDENT * 2),
              after: convertInchesToTwip(INDENT * 1),
            },
          },
        },
        {
          id: 'Text',
          name: 'Text',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            size: '16pt',
          },
        },
      ],
    },
    sections: [
      {
        children: children as Array<Paragraph>,
      },
    ],
  })

  switch (options.output) {
    case 'buffer':
      return await Packer.toBuffer(doc)
    case 'blob':
      return await Packer.toBlob(doc)
  }

  return await Packer.toBase64String(doc)

  function walk(node) {
    const children: Array<any> = []
    node.children.forEach(child => {
      children.push(...walkChild(child))
    })
    return children
  }

  // Override function
  function walkChild(node) {
    // console.log(node)
    const children: Array<any> = []
    switch (node.type) {
      case 'break':
        children.push({ type: 'paragraph', children: [] })
        break
      case 'list':
        children.push(
          ...walkList(node, { depth: 0, ordered: node.ordered }),
        )
        break
      case 'heading':
        children.push(...walkHeading(node))
        break
      case 'paragraph': {
        children.push(...walkParagraph(node))
        break
      }
    }
    return children
  }

  function walkParagraph(node) {
    const children: Array<any> = []
    node.children.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push({ type: 'text', text: child.value })
          break
        case 'link':
          children.push(
            walkLink(child, { italics: false, bold: false }),
          )
          break
        case 'emphasis':
          children.push(...walkEm(child, { bold: false }))
          break
        case 'strong':
          children.push(...walkStrong(child, { italics: false }))
          break
      }
    })
    return [
      { type: 'paragraph', children, style: 'Text' },
      { type: 'paragraph', children: [] },
    ]
  }

  function walkHeading(node) {
    const heading = HEADING[node.depth]
    const children: Array<any> = []

    node.children.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push({ type: 'text', text: child.value })
          break
        case 'link':
          children.push(
            walkLink(child, { italics: false, bold: false }),
          )
          break
        case 'emphasis':
          children.push(...walkEm(child, { bold: false }))
          break
        case 'strong':
          children.push(...walkStrong(child, { italics: false }))
          break
      }
    })

    return [
      { type: 'paragraph', heading, children },
      { type: 'paragraph', children: [] },
    ]
  }

  function walkStrong(node, { italics }) {
    const children: Array<any> = []

    node.children.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push({
            type: 'text',
            text: child.value,
            bold: true,
            italics,
          })
          break
        case 'link':
          children.push(walkLink(child, { italics, bold: true }))
          break
        case 'emphasis':
          children.push(...walkEm(child, { bold: true }))
          break
      }
    })

    return children
  }

  function walkEm(node, { bold }) {
    const children: Array<any> = []

    node.children.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push({
            type: 'text',
            text: child.value,
            bold,
            italics: true,
          })
          break
        case 'link':
          children.push(walkLink(child, { italics: true, bold }))
          break
        case 'strong':
          children.push(...walkStrong(child, { italics: true }))
          break
      }
    })

    return children
  }

  function walkLink(node, { italics, bold }) {
    const children: Array<any> = []
    node.children.forEach(child => {
      switch (child.type) {
        case 'text':
          children.push({
            type: 'text',
            text: child.value,
            italics,
            bold,
          })
          break
        case 'emphasis':
          children.push(...walkEm(child, { bold }))
          break
        case 'strong':
          children.push(...walkStrong(child, { italics }))
          break
      }
    })
    return { type: 'link', link: node.url, children }
  }

  function walkText(node, { italics, bold }) {
    const children: Array<any> = []
    if (node.children?.length) {
      node.children.forEach(child => {
        switch (child.type) {
          case 'text':
            children.push({ type: 'text', text: child.value })
            break
          case 'link':
            children.push(walkLink(child, { italics, bold }))
            break
          case 'emphasis':
            children.push(...walkEm(child, { bold }))
            break
          case 'strong':
            children.push(...walkStrong(child, { italics }))
            break
        }
      })
    } else {
      children.push({ type: 'text', text: node.value })
    }
    return children
  }

  function walkList(node, { ordered, depth = 0 }) {
    const children: Array<any> = []
    node.children.forEach(item => {
      children.push(...walkListItem(item, { ordered, depth }))
    })
    if (depth === 0) {
      children.push({ type: 'paragraph', children: [] })
    }
    return children
  }

  function walkListItem(node, { ordered, depth = 0 }) {
    const opts = ordered
      ? {
          numbering: {
            reference: `number`,
            level: depth,
          },
        }
      : {
          numbering: {
            reference: `bullet`,
            level: depth,
          },
        }
    // : {
    //     bullet: {
    //       level: depth,
    //     },
    //   }
    const items: Array<any> = []
    const children: Array<any> = []

    node.children.forEach(child => {
      switch (child.type) {
        case 'paragraph':
          children.push(
            ...walkText(child, { italics: false, bold: false }),
          )
          break
        case 'list':
          items.push(
            ...walkList(child, {
              ordered: child.ordered,
              depth: depth + 1,
            }),
          )
          break
      }
    })

    items.unshift({
      type: 'paragraph',
      ...opts,
      children,
      style: 'Text',
    })
    return items
  }

  function convert(child) {
    switch (child.type) {
      case 'text':
        return new TextRun({
          ...child,
        })
      case 'paragraph':
        return new Paragraph({
          ...child,
          children: child.children.map(convert),
        })
      case 'run':
        return new Run(child)
      case 'link':
        return new ExternalHyperlink({
          ...child,
          children: child.children.map(convert),
        })
    }
  }
}
