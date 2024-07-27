# results/models.py

from django.db import models

class Student(models.Model):
    student_id = models.CharField(max_length=50)
    name = models.CharField(max_length=100)
    result = models.CharField(max_length=20)

    def __str__(self):
        return self.name
