# -*- coding: utf-8 -*-

import pkg_resources


import hashlib
import logging
import mimetypes
import os
import uuid

# from docx import Document

from xblock.core import XBlock
from xblock.fields import Scope, Integer, String, JSONField
from xblock.fragment import Fragment
from django.core.files import File
from django.core.files.storage import default_storage
from django.core.exceptions import PermissionDenied
from django.views.decorators.csrf import csrf_exempt

from student.models import user_by_anonymous_id
from submissions import api as submissions_api
from submissions.models import StudentItem as SubmissionsStudent

from functools import partial
from xmodule.util.duedate import get_extended_due_date
from webob.response import Response


from .utils import (
    load_resource,
    render_template,
    load_resources,
    )

class XlsxCheckerXBlock(XBlock):
    """
    TO-DO: document what your XBlock does.
    """
    lab_scenario = Integer(
        display_name=u"Номер сценария",
        help=(u"Номер сценария",
              u"Номер сценария"),
        default=0,
        scope=Scope.settings
    )

    lines_settings = JSONField(
        display_name=u"Настройки сценария",
        help=u"Настройки сценария",
        default={
            "lab1": {"instruction_name": "instruction_lab_1.docx", "number": 1, "name": "Формулы, функции и диаграммы в процессоре Microsoft Office Excel"},
            "lab2": {"instruction_name": "instruction_lab_2.docx", "number": 2, "name": "Построение графиков функций"},
            "lab3": {"instruction_name": "instruction_lab_3.docx", "number": 3, "name": "Сортировка, фильтры и промежуточные итоги"},
        },
        scope=Scope.settings
    )

    correct_xlsx_uid = String(
         default='', scope=Scope.settings,
         help='Correct file from teacher',
        )
    correct_xlsx_name = String(
         default='', scope=Scope.settings,
         help='Name of correct file from teacher',
        )

    student_xlsx_uid = String(
         default='', scope=Scope.user_state,
         help='Studen file from student',
        )
    student_xlsx_name = String(
         default='', scope=Scope.user_state,
         help='Name of student file from student',
        )

    display_name = String(
        display_name=u"Название",
        help=u"Название задания, которое увидят студенты.",
        default=u'Проверка MS Excel',
        scope=Scope.settings
    )

    question = String(
        # TODO: list
        display_name=u"Вопрос",
        help=u"Текст задания.",
        default=u"Лабораторные работы MS Excel",
        scope=Scope.settings
    )

    weight = Integer(
        display_name=u"Максимальное количество баллов",
        help=(u"Максимальное количество баллов",
              u"которое может получить студент."),
        default=10,
        scope=Scope.settings
    )

    #TODO: 1!
    max_attempts = Integer(
        display_name=u"Максимальное количество попыток",
        help=u"",
        default=10,
        scope=Scope.settings
    )
    
    attempts = Integer(
        display_name=u"Количество использованных попыток",
        help=u"",
        default=0,
        scope=Scope.user_state
    )

    points = Integer(
        display_name=u"Текущее количество баллов студента",
        default=None,
        scope=Scope.user_state
    )

    def resource_string(self, path):
        """Handy helper for getting resources from our kit."""
        data = pkg_resources.resource_string(__name__, path)
        return data.decode("utf8")

    # TO-DO: change this view to display your data your own way.
    def student_view(self, context=None):
        context = {
            "display_name": self.display_name,
            "weight": self.weight,
            "question": self.question,
            "student_xlsx_name": self.student_xlsx_name,
            "points": self.points,
            "attempts": self.attempts,
        }

        if self.max_attempts != 0:
            context["max_attempts"] = self.max_attempts

        if self.past_due():
            context["past_due"] = True

        if answer_opportunity(self):
            context["answer_opportunity"] = True

        fragment = Fragment()
        fragment.add_content(
            render_template(
                "static/html/xlsx_checker.html",
                context
            )
        )

        js_urls = (
            "static/js/src/dxlsx_checker.js",
            )

        css_urls = (
            "static/css/xlsx_checker.css",
            )

        load_resources(js_urls, css_urls, fragment)

        fragment.initialize_js('XlsxCheckerXBlock')
        return fragment

    def studio_view(self, context=None):
        context = {
            "display_name": self.display_name,
            "weight": self.weight,
            "question": self.question,
            "max_attempts": self.max_attempts,
        }

        fragment = Fragment()
        fragment.add_content(
            render_template(
                "static/html/xlsx_checker_studio.html",
                context
            )
        )

        js_urls = (
            "static/js/src/xlsx_checker_studio.js",
            )

        css_urls = (
            "static/css/xlsx_checker_studio.css",
            )

        load_resources(js_urls, css_urls, fragment)

        fragment.initialize_js('XlsxCheckerXBlock')
        return fragment

    # TO-DO: change this handler to perform your own actions.  You may need more
    # than one handler, or you may not need any handlers at all.
    @XBlock.json_handler
    def increment_count(self, data, suffix=''):
        """
        An example handler, which increments the data.
        """
        # Just to show data coming in...
        assert data['hello'] == 'world'

        self.count += 1
        return {"count": self.count}

    # TO-DO: change this to create the scenarios you'd like to see in the
    # workbench while developing your XBlock.
    @staticmethod
    def workbench_scenarios():
        """A canned scenario for display in the workbench."""
        return [
            ("XlsxCheckerXBlock",
             """<xlsx_checker/>
             """),
            ("Multiple XlsxCheckerXBlock",
             """<vertical_demo>
                <xlsx_checker/>
                <xlsx_checker/>
                <xlsx_checker/>
                </vertical_demo>
             """),
        ]

    @XBlock.json_handler
    def student_submit(self, data, suffix=''):

        def check_answer():
            return 55

        grade_global = check_answer()
        self.points = grade_global
        self.points = grade_global * self.weight / 100
        self.points = int(round(self.points))
        self.attempts += 1
        self.runtime.publish(self, 'grade', {
            'value': self.points,
            'max_value': self.weight,
        })
        res = {"success_status": 'ok', "points": self.points, "weight": self.weight, "attempts": self.attempts, "max_attempts": self.max_attempts}
        return res

    @XBlock.json_handler
    def studio_submit(self, data, suffix=''):
        self.display_name = data.get('display_name')
        self.question = data.get('question')
        self.weight = data.get('weight')
        self.max_attempts = data.get('max_attempts')

        return {'result': 'success'}

    @XBlock.handler
    def download_student_file(self, request, suffix=''):
        path = self._students_storage_path(self.student_xlsx_uid, self.student_xlsx_name)
        return self.download(
            path,
            mimetypes.guess_type(self.source_xlsx_name)[0],
            self.student_xlsx_name
        )


    def is_course_staff(self):
        # pylint: disable=no-member
        """
         Check if user is course staff.
        """
        return getattr(self.xmodule_runtime, 'user_is_staff', False)

    @XBlock.handler
    def student_filename(self, request, suffix=''):
        return Response(json_body={'student_filename': self.student_xlsx_name})

    @XBlock.handler
    def upload_student_file(self, request, suffix=''):
        upload = request.params['studentFile']
        self.student_xlsx_name = upload.file.name
        self.student_xlsx_uid = uuid.uuid4().hex
        path = self._students_storage_path(self.student_xlsx_uid, self.student_xlsx_name)
        if not default_storage.exists(path):
            default_storage.save(path, File(upload.file))
        obj = path
        return Response(json_body=obj)

    def _students_storage_path(self, uid, filename):
        # pylint: disable=no-member
        """
        Get file path of storage.
        """
        path = (
            '{loc.org}/{loc.course}/{loc.block_type}/students'
            '/{uid}{ext}'.format(
                loc=self.location,
                uid=uid,
                ext=os.path.splitext(filename)[1]
            )
        )
        return path

    def download(self, path, mime_type, filename, require_staff=False):
        """
        Return a file from storage and return in a Response.
        """
        try:
            file_descriptor = default_storage.open(path)
            app_iter = iter(partial(file_descriptor.read, BLOCK_SIZE), '')
            return Response(
                app_iter=app_iter,
                content_type=mime_type,
                content_disposition="attachment; filename=" + filename.encode('utf-8'))
        except IOError:
            if require_staff:
                return Response(
                    "Sorry, assignment {} cannot be found at"
                    " {}. Please contact {}".format(
                        filename.encode('utf-8'), path, settings.TECH_SUPPORT_EMAIL
                    ),
                    status_code=404
                )
            return Response(
                "Sorry, the file you uploaded, {}, cannot be"
                " found. Please try uploading it again or contact"
                " course staff".format(filename.encode('utf-8')),
                status_code=404
            )

    def past_due(self):
            """
            Проверка, истекла ли дата для выполнения задания.
            """
            due = get_extended_due_date(self)
            if due is not None:
                if _now() > due:
                    return False
            return True

    def is_course_staff(self):
        """
        Проверка, является ли пользователь автором курса.
        """
        return getattr(self.xmodule_runtime, 'user_is_staff', False)

    def is_instructor(self):
        """
        Проверка, является ли пользователь инструктором.
        """
        return self.xmodule_runtime.get_user_role() == 'instructor'

def _now():
    """
    Получение текущих даты и времени.
    """
    return datetime.datetime.utcnow().replace(tzinfo=pytz.utc)

def answer_opportunity(self):
    """
    Возможность ответа (если количество сделанных попыток меньше заданного).
    """
    if self.max_attempts <= self.attempts and self.max_attempts != 0:
        return False
    else:
        return True

def require(assertion):
    """
    Raises PermissionDenied if assertion is not true.
    """
    if not assertion:
        raise PermissionDenied
