function now() {
    return new Date().toLocaleDateString()
}

function nowPlus20Days() {
    var date = new Date()
    date.setDate(date.getDate() + 20);
    return date.toLocaleDateString();
}

function total(items) {
    var sum = 0
    //if (!items) return sum;
    items.forEach(function (i) {
        console.log('Calculating item ' + i.name + '; you should see this message in debug run')
        sum += i.price
    })
    return sum
}

function getPageNumber(pageIndex) {
    if (pageIndex == null) {
        return ''
    }

    const pageNumber = pageIndex + 1

    return pageNumber
}

function getTotalPages(pages) {
    if (!pages) {
        return ''
    }

    return pages.length
}

/*
function getPageNumber(pages, pageIndex) {
    if (!pages || pageIndex == null) {
        return ''
    }

    const pagesToIgnore = pages.reduce((acu, page) => {
        const shouldIgnore = page.items.find((p) => p.ignorePageInCount === true) != null

        if (shouldIgnore) {
            acu.push(page)
        }

        return acu
    }, []).length

    const pageNumber = pageIndex + 1

    return pageNumber - pagesToIgnore
}

function getTotalPages(pages) {
    if (!pages) {
        return ''
    }

    const pagesToIgnore = pages.reduce((acu, page) => {
        const shouldIgnore = page.items.find((p) => p.ignorePageInCount === true) != null

        if (shouldIgnore) {
            acu.push(page)
        }

        return acu
    }, []).length

    return pages.length - pagesToIgnore
}

*/