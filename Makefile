validate:
	@./scripts/validate.sh

test:
	cargo test -p deckmint

check: test validate
