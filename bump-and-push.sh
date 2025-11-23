#!/bin/bash
set -e

# ----- CONFIG -----
VERSION_FILE=".version"
TAG_PREFIX="v"
BRANCH="$(git rev-parse --abbrev-ref HEAD)"

# ----- CHECK GIT -----
if ! git rev-parse --is-inside-work-tree >/dev/null 2>&1; then
    echo "❌ Not inside a Git repository!"
    exit 1
fi

if [ "$BRANCH" = "HEAD" ]; then
    echo "❌ You are in a detached HEAD state — aborting."
    exit 1
fi

# ----- READ VERSION FILE -----
if [ ! -f "$VERSION_FILE" ]; then
    echo "0.0.0" > "$VERSION_FILE"
fi

OLD_VERSION=$(cat "$VERSION_FILE")

IFS='.' read -r MAJ MIN PATCH <<< "$OLD_VERSION"

PATCH=$((PATCH + 1))
NEW_VERSION="$MAJ.$MIN.$PATCH"

echo "$NEW_VERSION" > "$VERSION_FILE"

echo "🔢 Version bump: $OLD_VERSION → $NEW_VERSION"

# ----- COMMIT -----
git add "$VERSION_FILE"
git commit -m "Bump version to $NEW_VERSION" || true

# ----- TAG -----
TAG="${TAG_PREFIX}${NEW_VERSION}"

echo "🏷️  Creating/Updating tag: $TAG"
git tag -f "$TAG"

# ----- PUSH -----
echo "🚀 Pushing to origin..."
git push origin "$BRANCH"
git push origin "$TAG" --force

echo "✅ Done. Version is now $NEW_VERSION"
