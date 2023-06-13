# ExcelApp

```html
GET http://localhost:8080/search
```
This endpoint renders the "search" view.

### Request Parameters

None

### Response

View name: "search"

Model attributes: None

```html
POST http://localhost:8080/search
```
This endpoint performs a search based on the provided query parameter and renders the "searchResults" view.

### Request Parameters

query (required): The search query.

### Response

View name: "searchResults"

Model attributes:

"results": The Excel sheet containing the search results (if any).

"error": An error message if no results were found.
