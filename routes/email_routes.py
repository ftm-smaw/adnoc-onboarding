from flask import Blueprint, jsonify
from services.email_fetcher import fetch_and_save_attachments

email_bp = Blueprint('email', __name__)

@email_bp.route('/fetch-replies', methods=['POST'])
def fetch_replies():
    try:
        fetch_and_save_attachments()
        return jsonify({"status": "success", "message": "Attachments downloaded and renamed"})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})