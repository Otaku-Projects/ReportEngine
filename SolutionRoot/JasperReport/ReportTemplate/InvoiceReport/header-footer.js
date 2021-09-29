function getPageNumberByIndex(pageIndex) {
	if (pageIndex == null) {
		return ''
	}

	const pageNumber = pageIndex + 1

	return pageNumber
}

function getPageNumber(pageNumber) {
	if (typeof (pageNumber) == "undefined" || pageNumber == null) {
		return '0';
	}

	return pageNumber
}

function getTotalPages(pages) {
	if (!pages) {
		return ''
	}

	return pages.length
}

function showIfOnFirstPage(pageNumber, options) {
	//var pageNumber = parseInt(pageIndex) + 1;

	//if (pageIndex === null)
	//	return options.inverse(this);

	if (pageNumber > 1) {
		return options.fn(this);
	} else {
		return options.inverse(this);
	}

	//return pageNumber;
}