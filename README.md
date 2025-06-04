I (well mostly ChatGPT) made a script that will download and index all your Zendesk tickets and generate an HTML file for offline viewing, using Zendesk's API.

A customer wanted to avoid upgrading their Zendesk plan, and was wondering if they could just download all their ticket data in case they would need it in the future. I did some playing around with ChatGPT and came up with this.

What you need to know to make this work:

1. Install wkhtmltopdf. This is used to generate PDFs of the tickets for easy printing. This can be disabled in the script if you wish.

2. Be sure to set the variables at the beginning of the script, especially "$subomain", "$email", and "$api_token". Optionally, you can change "$ticketsFolder" to save the data somewhere else. By default it will save everything wherever the script is located. You can also set "generatepdfs" to "$false" if you wish (makes the process a bit quicker), and if you installed "wkhtmltopdf" in a non-default location, be sure to update this as well.

3. Feel free to make your own tweaks and improvements. At this point, it just downloads everything, with no way to filter by date or anything. I'm sure it's possible to add this functionality, but I didn't take the time to do it because my customer doesn't care about that.
