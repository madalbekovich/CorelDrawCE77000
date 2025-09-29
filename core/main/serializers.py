from rest_framework import serializers
from .models import EventsHandler

class EventsHandlerSerializer(serializers.ModelSerializer):
    class Meta:
        model = EventsHandler
        fields = "__all__"
