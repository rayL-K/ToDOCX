#!/bin/sh

set -eu

PROTECTED_BRANCHES="main master develop"
TYPE_PATTERN='(feature|bugfix|hotfix|release|refactor|docs|chore|test|experiment|spike|wip)'
BRANCH_PATTERN="^${TYPE_PATTERN}/(([A-Za-z]+-[0-9]+-)?[a-z0-9]+(-[a-z0-9]+)*)$"
MAX_BRANCH_LENGTH=60

print_error() {
	printf '%s\n' "$@" >&2
}

current_branch() {
	git symbolic-ref --quiet --short HEAD 2>/dev/null || return 1
}

is_protected_branch() {
	case " $PROTECTED_BRANCHES " in
		*" $1 "*) return 0 ;;
		*) return 1 ;;
	esac
}

validate_named_branch() {
	branch="$1"

	if [ "${#branch}" -gt "$MAX_BRANCH_LENGTH" ]; then
		print_error "分支名过长: $branch"
		print_error "请控制在 ${MAX_BRANCH_LENGTH} 个字符以内。"
		return 1
	fi

	if printf '%s' "$branch" | grep -Eq "$BRANCH_PATTERN"; then
		return 0
	fi

	print_error "分支名不符合规范: $branch"
	print_error "允许格式:"
	print_error "  <type>/<kebab-case-description>"
	print_error "  <type>/<ticket-id>-<kebab-case-description>"
	print_error "允许的 type: feature, bugfix, hotfix, release, refactor, docs, chore, test, experiment, spike, wip"
	print_error "示例: feature/add-dark-mode-toggle"
	print_error "示例: bugfix/JIRA-123-fix-login-timeout"
	return 1
}

validate_pre_commit() {
	branch="$(current_branch || exit 0)"

	if is_protected_branch "$branch"; then
		exit 0
	fi

	validate_named_branch "$branch"
}

validate_pre_push() {
	local_ref="${1:-}"

	if [ -n "$local_ref" ] && [ "${local_ref#refs/heads/}" != "$local_ref" ]; then
		branch="${local_ref#refs/heads/}"
	else
		branch="$(current_branch || exit 0)"
	fi

	[ -n "$branch" ] || exit 0

	# 当前仓库允许直接 push 到 main，但仍保留普通主题分支命名校验。
	if is_protected_branch "$branch"; then
		return 0
	fi

	validate_named_branch "$branch"
}

case "${1:-}" in
	validate_pre_commit)
		validate_pre_commit
		;;
	validate_pre_push)
		validate_pre_push "${2:-}"
		;;
	*)
		print_error "未知 hook 命令: ${1:-}"
		exit 2
		;;
esac
