# vqsfia
Form generation for SFIA-based progression

## Quick start

### Form generation to files

Clone the repo.

Run `npm install`

See the shell scripts for examples on how to create a skills profile and a progression matrix.

Run `node index.js -h` for help on command line options syntax.


### As a web server

Run `node index.js -s 3000` where 3000 is the port you'd like to serve on.

### Hacking

To build the client, you'll need to change into the `client` directory and:

Run `npm install`
Run `webpack` (installing first [following instructions](webpack.github.io) if necessary)
