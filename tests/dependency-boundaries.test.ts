import { readdirSync, statSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import * as ts from 'typescript'
import { describe, expect, it } from 'vitest'

type ModuleReference = {
  sourceFile: string
  specifier: string
}

type ResolvedModule =
  | {
      kind: 'external'
      packageName: string
      specifier: string
    }
  | {
      kind: 'local'
      projectPath: string
      specifier: string
    }

type Restriction = {
  description: string
  matches: (module: ResolvedModule) => boolean
}

type BoundaryRule = {
  layer: string
  restrictions: Restriction[]
  sourceMatches: (projectPath: string) => boolean
}

type BoundaryViolation = {
  description: string
  layer: string
  sourcePath: string
  specifier: string
}

const projectRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const sourceRoot = path.join(projectRoot, 'src')

function toPosix(value: string) {
  return value.split(path.sep).join('/')
}

function toProjectPath(absolutePath: string) {
  return toPosix(path.relative(projectRoot, absolutePath))
}

function listSourceFiles(directory: string): string[] {
  return readdirSync(directory).flatMap((entry) => {
    const absolutePath = path.join(directory, entry)
    const stats = statSync(absolutePath)

    if (stats.isDirectory()) {
      return listSourceFiles(absolutePath)
    }

    if (/\.(tsx?|mts|cts)$/.test(entry) && !entry.endsWith('.d.ts')) {
      return [absolutePath]
    }

    return []
  })
}

function getPackageName(specifier: string) {
  if (specifier.startsWith('@')) {
    return specifier.split('/').slice(0, 2).join('/')
  }

  return specifier.split('/')[0]
}

function getLocalImportPath(sourceFile: string, specifier: string) {
  if (specifier.startsWith('.')) {
    return path.resolve(path.dirname(sourceFile), specifier)
  }

  if (specifier.startsWith('/src/')) {
    return path.join(projectRoot, specifier.slice(1))
  }

  if (specifier.startsWith('src/')) {
    return path.join(projectRoot, specifier)
  }

  if (specifier.startsWith('@/')) {
    return path.join(sourceRoot, specifier.slice(2))
  }

  return null
}

function resolveModuleReference(reference: ModuleReference): ResolvedModule {
  const localImportPath = getLocalImportPath(reference.sourceFile, reference.specifier)

  if (localImportPath) {
    return {
      kind: 'local',
      projectPath: toProjectPath(localImportPath),
      specifier: reference.specifier,
    }
  }

  return {
    kind: 'external',
    packageName: getPackageName(reference.specifier),
    specifier: reference.specifier,
  }
}

function getModuleReferences(sourceFile: string): ModuleReference[] {
  const sourceText = ts.sys.readFile(sourceFile)
  if (!sourceText) {
    return []
  }

  const sourceFileNode = ts.createSourceFile(sourceFile, sourceText, ts.ScriptTarget.Latest, true)
  const references: ModuleReference[] = []

  function addStringModuleSpecifier(node: ts.Node) {
    if (ts.isStringLiteralLike(node)) {
      references.push({
        sourceFile,
        specifier: node.text,
      })
    }
  }

  function visit(node: ts.Node) {
    if ((ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) && node.moduleSpecifier) {
      addStringModuleSpecifier(node.moduleSpecifier)
    }

    if (
      ts.isImportTypeNode(node) &&
      ts.isLiteralTypeNode(node.argument) &&
      ts.isStringLiteralLike(node.argument.literal)
    ) {
      addStringModuleSpecifier(node.argument.literal)
    }

    if (ts.isCallExpression(node) && node.expression.kind === ts.SyntaxKind.ImportKeyword) {
      const [firstArgument] = node.arguments
      if (firstArgument) {
        addStringModuleSpecifier(firstArgument)
      }
    }

    ts.forEachChild(node, visit)
  }

  visit(sourceFileNode)
  return references
}

function importsPackage(packageName: string) {
  return (module: ResolvedModule) => module.kind === 'external' && module.packageName === packageName
}

function importsPackagePrefix(packagePrefix: string) {
  return (module: ResolvedModule) => module.kind === 'external' && module.specifier.startsWith(packagePrefix)
}

function importsAnyPackage(...packageNames: string[]) {
  return (module: ResolvedModule) => packageNames.some((packageName) => importsPackage(packageName)(module))
}

function importsLocalPath(localPath: string) {
  return (module: ResolvedModule) =>
    module.kind === 'local' && (module.projectPath === localPath || module.projectPath.startsWith(`${localPath}/`))
}

function importsLocalPathPrefix(localPathPrefix: string) {
  return (module: ResolvedModule) => module.kind === 'local' && module.projectPath.startsWith(localPathPrefix)
}

function importsApplicationOutsidePorts(module: ResolvedModule) {
  return (
    module.kind === 'local' &&
    importsLocalPath('src/application')(module) &&
    !importsLocalPath('src/application/ports')(module)
  )
}

const importsReact = importsAnyPackage('react', 'react-dom')
const importsExcelJs = importsPackage('exceljs')
const importsVite = (module: ResolvedModule) =>
  importsPackage('vite')(module) || importsPackagePrefix('@vitejs/')(module)

const boundaryRules: BoundaryRule[] = [
  {
    layer: 'domain',
    sourceMatches: importsLocalPath('src/domain'),
    restrictions: [
      { description: 'React packages', matches: importsReact },
      { description: 'ExcelJS package', matches: importsExcelJs },
      { description: 'Vite packages', matches: importsVite },
      { description: 'application layer', matches: importsLocalPath('src/application') },
      { description: 'infrastructure layer', matches: importsLocalPath('src/infrastructure') },
      { description: 'adapter layer', matches: importsLocalPath('src/adapters') },
    ],
  },
  {
    layer: 'application',
    sourceMatches: importsLocalPath('src/application'),
    restrictions: [
      { description: 'React packages', matches: importsReact },
      { description: 'ExcelJS package', matches: importsExcelJs },
      { description: 'Vite packages', matches: importsVite },
      { description: 'infrastructure implementations', matches: importsLocalPath('src/infrastructure') },
      { description: 'adapter components', matches: importsLocalPath('src/adapters') },
      { description: 'browser entrypoint', matches: importsLocalPathPrefix('src/main') },
    ],
  },
  {
    layer: 'infrastructure',
    sourceMatches: importsLocalPath('src/infrastructure'),
    restrictions: [
      { description: 'React packages', matches: importsReact },
      { description: 'application services', matches: importsApplicationOutsidePorts },
      { description: 'React adapters', matches: importsLocalPath('src/adapters/react') },
    ],
  },
  {
    layer: 'React adapters',
    sourceMatches: importsLocalPath('src/adapters/react'),
    restrictions: [{ description: 'direct ExcelJS package imports', matches: importsExcelJs }],
  },
]

function findBoundaryViolations() {
  return listSourceFiles(sourceRoot).flatMap((sourceFile): BoundaryViolation[] => {
    const sourcePath = toProjectPath(sourceFile)
    const rulesForSource = boundaryRules.filter((rule) =>
      rule.sourceMatches({
        kind: 'local',
        projectPath: sourcePath,
        specifier: sourcePath,
      }),
    )

    return getModuleReferences(sourceFile).flatMap((reference) => {
      const resolvedModule = resolveModuleReference(reference)

      return rulesForSource.flatMap((rule) =>
        rule.restrictions
          .filter((restriction) => restriction.matches(resolvedModule))
          .map((restriction) => ({
            description: restriction.description,
            layer: rule.layer,
            sourcePath,
            specifier: reference.specifier,
          })),
      )
    })
  })
}

function formatViolations(violations: BoundaryViolation[]) {
  return violations
    .map(
      (violation) =>
        `${violation.sourcePath} imports ${violation.specifier} (${violation.layer} cannot depend on ${violation.description})`,
    )
    .join('\n')
}

describe('dependency boundaries', () => {
  it('keeps source imports within onion architecture layers', () => {
    const violations = findBoundaryViolations()

    expect(violations, formatViolations(violations)).toEqual([])
  })
})
