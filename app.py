from flask import Flask, request, jsonify, render_template, redirect, url_for

app = Flask(__name__)

@app.route("/")
def main():
    print("Route hit!")
    return render_template("main.html")

if __name__ == "__main__":
    app.run(debug=True)
