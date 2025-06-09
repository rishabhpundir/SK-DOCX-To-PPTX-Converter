from django.urls import path
from . import views

app_name = 'converter'

urlpatterns = [
    # Home page with upload form
    path('', views.HomeView.as_view(), name='home'),
    
    # Download page
    path('download/<int:pk>/', views.DownloadView.as_view(), name='download'),
    
    # Actual file download
    path('download/<int:pk>/file/', views.DownloadFileView.as_view(), name='download_file'),
    
    # Status check (for AJAX if needed)
    path('status/<int:pk>/', views.StatusView.as_view(), name='status'),
]