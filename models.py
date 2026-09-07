from datetime import datetime, timezone
from extensions import db

class UnlistedServiceRequest(db.Model):
    __tablename__ = 'unlisted_service_request'
    
    id = db.Column(db.Integer, primary_key=True)
    service_name = db.Column(db.String(255), nullable=False)
    status = db.Column(db.String(50), default='Pending')  # Pending, Approved, Added
    date_submitted = db.Column(db.DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

    def __repr__(self):
        return f'<UnlistedServiceRequest {self.service_name}>'


class Feedback(db.Model):
    __tablename__ = 'feedback'
    
    id = db.Column(db.Integer, primary_key=True)
    service = db.Column(db.String(100))
    sub_service = db.Column(db.String(100))
    rating = db.Column(db.String(50))
    comment = db.Column(db.Text)
    date_submitted = db.Column(db.DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))

    def __repr__(self):
        return f'<Feedback {self.id}>'