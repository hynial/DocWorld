#!/usr/local/bin/bash

CURRENTDIR=$(dirname "$0")
archiveName=$(basename $CURRENTDIR)
VERSION='1.0'
CURBRANCH=$(git branch | awk '/\*/ { print $2; }')
RELEASETAG=$CURBRANCH

if [ -f ./.gitattributes ]; then
  echo "exist .gitattributes file, make sure that been committed before git archive"
else
  # shellcheck disable=SC2129
  echo "/test export-ignore" >> .gitattributes
  echo ".gitattributes export-ignore" >> .gitattributes
  echo ".gitignore export-ignore" >> .gitattributes
  echo ".github export-ignore" >> .gitattributes
  echo ".vscode export-ignore" >> .gitattributes

  echo "generated git attribute file, please add&commit it! .e.g: git commit -m 'add/modify .gitattributes'"

  git add ./.gitattributes
  exit 0
fi

# modified | excluded
# git ls-files --modified -o --exclude-standard | grep .gitattributes
# new added
# git diff --diff-filter=A --name-only HEAD

gitatt=$( (git ls-files --modified -o --exclude-standard ; git diff --diff-filter=A --name-only HEAD) | grep .gitattributes )

if [ "$gitatt" = ".gitattributes" ]; then
  echo "Please commit '.gitattributes' before archive!"
  echo "e.g: git add .gitattributes & git commit -m 'add/modify .gitattributes'"
  exit 1
fi

echo "ready to archive current branch ..."
# tarball
# git archive --format=tar --prefix=${archiveName}-${VERSION}/ ${RELEASETAG} | bzip2 -9 > ${CURRENTDIR}/${archiveName}-${RELEASETAG}.tar.bz2
# zip ball
git archive --format zip --output ${CURRENTDIR}/${archiveName}-${RELEASETAG}-${VERSION}.zip ${RELEASETAG}

echo "finished archive action."
