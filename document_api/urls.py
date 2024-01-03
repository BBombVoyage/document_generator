from django.urls import path
from .views import generate_document

urlpatterns = [
    path('generate_document/', generate_document, name='generate_document'),
]