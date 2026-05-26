.PHONY: run push release retag

run:
	python drive_parser.py

push:
	git push origin main

release:
	@test -n "$(VERSION)" || (echo "Usage: make release VERSION=v1.2.3"; exit 1)
	git tag $(VERSION)
	git push origin main --tags

retag:
	@test -n "$(VERSION)" || (echo "Usage: make retag VERSION=v1.2.3"; exit 1)
	git tag -d $(VERSION) 2>/dev/null || true
	git push origin :$(VERSION) 2>/dev/null || true
	git tag $(VERSION)
	git push origin --tags
