from excel_git_viewer.task_coordinator import TaskCoordinator


def test_starting_a_new_task_cancels_and_invalidates_the_previous_task() -> None:
    coordinator = TaskCoordinator()

    first = coordinator.begin()
    second = coordinator.begin()

    assert first.cancellation.is_cancelled is True
    assert coordinator.is_current(first.task_id) is False
    assert coordinator.is_current(second.task_id) is True
