{% extends 'display_priority.tpl' %}


{% block in_prompt %}
{% endblock in_prompt %}

{% block output_prompt %}
{%- endblock output_prompt %}

{% block input %}
{{ resources["worker"].process_input(cell)}}
{% endblock input %}

{% block pyerr %}
{{ super() }}
{% endblock pyerr %}

{% block traceback_line %}
{{ line | indent | strip_ansi }}
{% endblock traceback_line %}

{%- block error -%}
```
##OUTPUT
{{ "\n".join(output.traceback) | strip_ansi }}
```
{%- endblock error -%}

{% block execute_result %}

{% block data_priority scoped %}
{{ super() }}
{% endblock %}
{% endblock execute_result %}

{% block stream %}
```
##OUTPUT
{{ output.text }}
```
{% endblock stream %}

{% block data_svg %}
{{ resources["worker"].process_graph(output, "svg") }}
{% endblock data_svg %}

{% block data_png %}
{{ resources["worker"].process_graph(output, "png") }}
{% endblock data_png %}

{% block data_jpg %}
{{ resources["worker"].process_graph(output, "jpg") }}
{% endblock data_jpg %}

{% block data_latex %}
{{ output["data"]["text/latex"] }}
{% endblock data_latex %}

{% block data_html scoped %}
{{ resources["worker"].process_html(output["data"]["text/html"]) }}
{% endblock data_html %}

{% block data_markdown scoped %}
{{ output["data"]["text/markdown"] }}
{% endblock data_markdown %}

{% block data_text scoped %}
{{ resources["worker"].process_output(output)}}
{% endblock data_text %}

{% block markdowncell scoped %}
{{ resources["worker"].process_markdown(cell.source) }}
{% endblock markdowncell %}


{% block headingcell scoped %}
{{ '#' * cell.level }} {{ cell.source | replace('\n', ' ') }}
{% endblock headingcell %}

{% block unknowncell scoped %}
unknown type  {{ cell.type }}
{% endblock unknowncell %}